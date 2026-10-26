# IQC Nexus — Smart Reminder Hub & Audit-Grade PDF Engine Design

**Date:** 2026-10-27  
**Author:** Nguyễn Hải Đường (`hai.duong`)  
**Status:** Approved

---

## 1. Problem Statement & Goals

1. **Reminder Bottlenecks:**
   - Individual course reminder button (`exportRemindSingle('${s.id}')` in `Index.html`) was missing its client handler despite `exportPendingList` existing in `Code.js`.
   - `exportConsolidatedReminder` created temporary Google Spreadsheets in Drive on every click (`5–12s` latency + Drive Trash clutter).
   - Lack of instant 1-Click Copy formatted messages grouped by **Part** for Knox Messenger / Zalo / Teams.
   - Automated emails (`sendIncompleteReport`) listed flat rows without a Part-level summary table.

2. **PDF Export Bottlenecks:**
   - `exportCourseReport` in `Code.js` resolved `drive:FILE_ID` signatures sequentially inside a `.forEach()` loop (`N+1` Drive I/O bottleneck, taking 40–90s for 150+ users).
   - PDF layout lacked an Executive Summary table (Total / Signed / Exception / Pending / Completion Rate %), Part-level breakdown, QR verification badge, row page-break protection, and 3-tier ISO/Samsung Quality approval signature blocks.

---

## 2. Architecture & Components

### A. Smart Reminder Hub (`#modalReminderHub` in `Index.html` + `getReminderData` in `Code.js`)
- **Backend API (`getReminderData(courseId, adminToken)`)**:
  - Returns structured JSON of pending assignments (either for a specific `courseId` or across all active/overdue courses when `courseId === null`) in `<300ms` without creating any temporary Drive files.
  - Includes `partSummary` (grouped counts per Part) and `items` (`courseId`, `courseName`, `endDate`, `status`, `uid`, `name`, `com`, `group`, `part`).
- **On-Demand Email Trigger (`triggerManualReminderEmail(courseId, adminToken)`)**:
  - Allows Admin to immediately dispatch the HTML reminder email (grouped by Part) to `RECIPIENT_EMAILS` with 1 click.
- **Client Modal (`#modalReminderHub`)**:
  - **Part Filter Pills**: Click any Part badge (`Tất cả`, `IQC 1P (4)`, `IQC 2P (2)`, etc.) to filter the preview table and message generator.
  - **1-Click Copy Message**: Formats a clean, executive reminder text grouped by Part with user names, employee IDs, deadline, and Web App link.
  - **1-Click Instant Excel (`.xlsx` via `SheetJS` in `<0.2s`)**: Generates a 2-sheet workbook (`Tổng hợp theo Part` + `Danh sách chi tiết`) directly in browser RAM with zero Google Drive overhead.

### B. Parallel Batch Signature Resolver & Dual-Mode Audit PDF Engine
- **Batch Signature Resolver (`resolveSignaturesBatch(signVals)` in `Code.js`)**:
  - Checks `CacheService.getScriptCache()` for cached Base64 thumbnails (`SIG_B64_<fileId>`).
  - Fetches all uncached Drive file IDs in **parallel** using `UrlFetchApp.fetchAll(requests)` with OAuth Bearer token (`https://www.googleapis.com/drive/v3/files/<fileId>?alt=media`).
  - Stores fetched Base64 strings in `CacheService` (6 hours TTL) and returns the resolved map in `<1.5s`.
- **Audit-Grade Report Payload (`getCourseAuditReportData(courseID, adminToken)` & `exportCourseReport(courseID, adminToken)`)**:
  - Computes full course metrics: `total`, `signedCount`, `exceptionCount`, `pendingCount`, `completionRate`, and `partBreakdown`.
  - Builds an ISO/Samsung Audit-compliant HTML template featuring:
    - Header with IQC Nexus branding, Course ID, Validity Period, and QR Verification code.
    - Executive Summary KPI table & Part Breakdown.
    - Participant Signature table with `page-break-inside: avoid; break-inside: avoid`.
    - 3-Tier Approval Signature Block (`Người lập biểu` / `Xem xét (Part Leader)` / `Phê duyệt (IQC Manager)`).
  - Supports **Instant Vector Print / Save as PDF** (opens high-DPI vector print window in 1s) and **Server Direct `.pdf` Download**.
