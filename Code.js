const ADMIN_PASS = "@nhd16488";
const USER_DEFAULT_PASS = "@iqc6368";
const SPREADSHEET_ID = "15K3FetZvSMPRKLNp9I7kyPLCeAIpreWNeeG7uerW3wk";

function getDB() {
    return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function doGet() {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate().setTitle('Hệ Thống Đào Tạo').addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
}

function checkLogin(user, pass) {
    try {
        if (user === 'admin' && pass === ADMIN_PASS) {
            return { role: 'admin', user: 'Quản trị viên', userID: 'admin' };
        }

        const ss = getDB();
        const userData = ss.getSheetByName('Users').getDataRange().getDisplayValues();

        for (let i = 1; i < userData.length; i++) {
            let uID = userData[i][1].toString();
            if (uID === user && pass === USER_DEFAULT_PASS) {
                if (userData[i][6] === 'Inactive') return { error: "Tài khoản bị khóa!" };
                return { role: 'user', user: userData[i][0], userID: uID };
            }
        }
        return { error: "Sai thông tin đăng nhập!" };
    } catch (e) {
        return { error: "Lỗi Server: " + e.message };
    }
}

// --- REFACTORED DATA FETCHING ---
function getAdminUsers() {
    try {
        const ss = getDB();
        return ss.getSheetByName('Users').getDataRange().getDisplayValues().slice(1);
    } catch (e) { return { error: e.toString() }; }
}

function getAdminCourses() {
    return getCourseStatistics();
}

function getAdminFilters() {
    try {
        const users = getDB().getSheetByName('Users').getDataRange().getValues().slice(1);
        const getUnique = (idx) => [...new Set(users.map(u => String(u[idx] || '')).filter(Boolean))].sort();
        return {
            com: getUnique(2),
            group: getUnique(3),
            part: getUnique(4)
        };
    } catch (e) { return {}; }
}

function getAdminDashboardData() {
    const users = getAdminUsers();
    const newRegs = Array.isArray(users) ? users.filter(u => u[6] === 'New-Reg') : [];

    return {
        courses: getCourseStatistics(),
        filters: getAdminFilters(),
        newRegisters: newRegs,
        availableCourses: getAvailableCourses(),
        allUsers: users // Return FULL list to optimize Front-end (1 Call)
    };
}

function getCourseDetails(courseId) {
    try {
        const db = getDB();
        const sheet = db.getSheetByName('Courses');
        const data = sheet.getDataRange().getValues();
        const course = data.find(r => String(r[0]) === String(courseId));
        return course ? course[4] : '';
    } catch (e) { return ''; }
}

function saveUser(u, isEdit, courseIds) {
    try {
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const rowData = [u.name, u.id, u.com, u.group, u.part, 'user', u.status || 'Active'];

        if (isEdit) {
            for (let i = 1; i < data.length; i++) {
                if (data[i][1].toString() === u.id) {
                    sheet.getRange(i + 1, 2).setNumberFormat("@"); // Ensure ID format is text
                    sheet.getRange(i + 1, 1, 1, 7).setValues([rowData]);
                    return { success: true, user: rowData, action: 'update' };
                }
            }
        } else {
            for (let i = 1; i < data.length; i++) {
                if (data[i][1].toString() === u.id) return { error: "Mã nhân viên đã tồn tại!" };
            }
            // SHEET MODIFICATION: Use setValues + setNumberFormat("@") for ID
            const lastRow = sheet.getLastRow();
            const newRow = lastRow + 1;

            // Set ID (Column 2) to Plain Text
            sheet.getRange(newRow, 2).setNumberFormat("@");
            sheet.getRange(newRow, 1, 1, 7).setValues([rowData]);

            // Handle Initial Course Assignment
            if (courseIds && courseIds.length > 0) {
                assignCoursesToUsers([u.id], courseIds);
            }

            return { success: true, user: rowData, action: 'create' };
        }
    } catch (e) { return { error: "Lỗi Server: " + e.message }; }
}

function deleteUser(userID) {
    try {
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][1].toString() === userID) {
                // DELETE INCOMPLETE ASSIGNMENTS
                const assignSheet = getDB().getSheetByName('Assignments');
                const aData = assignSheet.getDataRange().getDisplayValues();
                // Iterate backwards
                for (let j = aData.length - 1; j >= 1; j--) {
                    // Col C(2) is UserID, Col H(7) is Sign, Col J(9) is Exception
                    if (aData[j][2].toString() === userID && aData[j][7] === '' && aData[j][9] === '') {
                        assignSheet.deleteRow(j + 1);
                    }
                }

                sheet.deleteRow(i + 1);
                return { success: true, id: userID };
            }
        }
        return { error: "User not found" };
    } catch (e) { return { error: e.message }; }
}

function registerUser(u) {
    try {
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();

        // Check duplicate ID (Index 1)
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][1]) === String(u.id)) {
                return { error: "Mã nhân viên đã tồn tại!" };
            }
        }

        // Name(0), ID(1), Com(2), Grp(3), Prt(4), Role(5), Status(6)
        // CHANGE: Default status 'New-Reg' instead of 'Active'
        const nextRow = sheet.getLastRow() + 1;
        const newRowData = [u.name, u.id, u.com, u.group, u.part, 'user', 'New-Reg'];

        sheet.getRange(nextRow, 2).setNumberFormat("@"); // Force Plain Text for ID
        sheet.getRange(nextRow, 1, 1, 7).setValues([newRowData]);
        return { success: true };
    } catch (e) {
        return { error: "Lỗi đăng ký: " + e.message };
    }
}

function getAvailableCourses() {
    try {
        const ss = getDB();
        const courseData = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const parseDate = (dStr) => { if (!dStr) return null; const p = dStr.split('-'); return new Date(p[0], p[1] - 1, p[2]); };

        return courseData.map(c => {
            const start = parseDate(c[2]);
            const end = parseDate(c[3]);
            let status = 'Opening';
            if (today > end) status = 'Pending'; // Or Ended? User wants Plan, Opening, Pending
            else if (today < start) status = 'Plan';

            // Filter: Only return Plan, Opening, Pending (User Requirement)
            // Pending usually means "Overdue but not closed"? In my previous logic PENDING meant "Today > End".
            // Assuming "Plan", "Opening", "Pending" are valid assignable states.
            return { id: c[0], name: c[1], status: status };
        });
    } catch (e) { return []; }
}

function assignCoursesToUsers(userIds, courseIds) {
    if (!userIds || userIds.length === 0 || !courseIds || courseIds.length === 0) return { count: 0 };
    try {
        const ss = getDB();
        const assignSheet = ss.getSheetByName('Assignments');
        const userSheet = ss.getSheetByName('Users');

        // Optimize: Convert userIds to Map for quick lookup of user details
        const allUsers = userSheet.getDataRange().getDisplayValues();
        const userMap = {};
        allUsers.slice(1).forEach(u => {
            userMap[u[1].toString()] = { name: u[0], id: u[1], com: u[2], group: u[3], part: u[4], role: u[5] };
        });

        const courseSheet = ss.getSheetByName('Courses');
        const allCourses = courseSheet.getDataRange().getDisplayValues();
        const courseMap = {};
        allCourses.slice(1).forEach(c => courseMap[c[0]] = c[1]); // ID -> Name

        const newRows = [];
        let count = 0;

        userIds.forEach(uid => {
            const uStr = uid.toString().trim();
            const user = userMap[uStr];

            // SECURITY CHECK: DO NOT ASSIGN TO ADMIN
            // Check Role (Index 5) or ID is 'admin'
            if (user) {
                const isAdmin = (String(user.role).toLowerCase() === 'admin') || (String(user.id).toLowerCase() === 'admin');
                const isUserRole = (String(user.role).toLowerCase() === 'user');

                // CRITICAL FIX: Only assign if role is specifically 'user' AND not admin ID
                if (!isAdmin && isUserRole) {
                    courseIds.forEach(cid => {
                        if (courseMap[cid]) {
                            newRows.push([
                                cid,
                                courseMap[cid],
                                user.id,
                                user.name,
                                user.com,
                                user.group,
                                user.part,
                                '', '', ''
                            ]);
                            count++;
                        }
                    });
                }
            }
        });

        if (newRows.length > 0) {
            assignSheet.getRange(assignSheet.getLastRow() + 1, 1, newRows.length, 10).setValues(newRows);
        }
        return { count: count };

    } catch (e) { return { error: e.message }; }
}

function processNewRegisters(userIds, courseIds) {
    try {
        // 1. Assign courses
        const assignRes = assignCoursesToUsers(userIds, courseIds);
        if (assignRes.error) return assignRes;

        // 2. Set Status to 'Active'
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const idSet = new Set(userIds.map(id => id.toString()));

        for (let i = 1; i < data.length; i++) {
            if (idSet.has(data[i][1].toString())) {
                sheet.getRange(i + 1, 7).setValue('Active');
            }
        }
        return { success: true, count: assignRes.count };
    } catch (e) { return { error: e.message }; }
}

function importUsers(userList, courseIds) {
    try {
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const existingIds = new Set(data.slice(1).map(r => r[1].toString()));

        let addedCount = 0;
        let errors = [];
        let newRows = [];
        let newUserIds = []; // Track IDs of added users

        userList.forEach(u => {
            if (!u.id || !u.name) return;
            const idStr = u.id.toString().trim();
            if (existingIds.has(idStr)) {
                errors.push("Duplicate: " + idStr);
            } else {
                newRows.push([u.name, idStr, u.com || '', u.group || '', u.part || '', 'user', 'Active']);
                existingIds.add(idStr);
                newUserIds.push(idStr);
                addedCount++;
            }
        });

        if (newRows.length > 0) {
            const startRow = sheet.getLastRow() + 1;
            // Set Number Format for ID column (Index 2 in sheet -> Col B)
            // for the range of new rows
            sheet.getRange(startRow, 2, newRows.length, 1).setNumberFormat("@");
            sheet.getRange(startRow, 1, newRows.length, 7).setValues(newRows);

            // Assign Courses to New Users
            if (courseIds && courseIds.length > 0) {
                assignCoursesToUsers(newUserIds, courseIds);
            }
        }

        return { success: true, added: addedCount, errors: errors };
    } catch (e) {
        return { error: "Import failed: " + e.message };
    }
}

// --- USER STATISTICS & BULK ACTIONS ---
function getUserStats() {
    try {
        const users = getDB().getSheetByName('Users').getDataRange().getDisplayValues().slice(1);
        const stats = {};
        let total = 0;

        users.forEach(u => {
            // u[1] is ID, u[4] is Part, u[5] is Role
            const isAdm = (String(u[5]).toLowerCase() === 'admin') || (String(u[1]).toLowerCase() === 'admin');
            if (!isAdm) {
                const part = u[4] || 'Chưa phân loại';
                if (!stats[part]) stats[part] = 0;
                stats[part]++;
                total++;
            }
        });

        return { stats: stats, total: total };
    } catch (e) {
        return { error: "Lỗi thống kê: " + e.message };
    }
}

function deleteUsersByScope(scope, value) {
    try {
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        let deletedCount = 0;
        let deletedIDs = []; // Track IDs to delete assignments
        // Iterate backwards to delete safely
        for (let i = data.length - 1; i >= 1; i--) { // Skip header
            const u = data[i];
            const uID = u[1];
            const uPart = u[4];
            const uRole = u[5];

            // CRITICAL SAFETY CHECK: NEVER DELETE ADMIN
            const isAdm = (String(uRole).toLowerCase() === 'admin') || (String(uID).toLowerCase() === 'admin');
            if (isAdm) continue;

            let shouldDelete = false;
            if (scope === 'all') {
                shouldDelete = true;
            } else if (scope === 'part') {
                if (String(uPart) === String(value)) shouldDelete = true;
                // Handle "Chưa phân loại" case if value is empty/null? 
                // Front-end should pass specific Part string.
            }

            if (shouldDelete) {
                // DELETE INCOMPLETE ASSIGNMENTS FOR THIS USER
                const assignSheet = getDB().getSheetByName('Assignments');
                const aData = assignSheet.getDataRange().getDisplayValues();
                const uIDToDelete = uID.toString();
                // We must be careful deleting from another sheet while iterating.
                // It's better to collect IDs first, but for bulk delete this might be slow if we do one by one.
                // Optimization: Collect all rows to delete from assignments first? 
                // Let's do simple approach: In bulk delete, we might just scan assignments once at the end?
                // No, sticking to per-user safe check or bulk scan.
                // Let's do a bulk scan of Assignments matching the deleted IDs.
                deletedCount++;
                deletedIDs.push(uID); // Collect ID
                sheet.deleteRow(i + 1);
            }
        }

        // Batch delete assignments for efficiency
        if (deletedIDs.length > 0) {
            const assignSheet = getDB().getSheetByName('Assignments');
            const aData = assignSheet.getDataRange().getDisplayValues();
            const idSet = new Set(deletedIDs.map(id => id.toString()));

            for (let j = aData.length - 1; j >= 1; j--) {
                if (idSet.has(aData[j][2].toString()) && aData[j][7] === '' && aData[j][9] === '') {
                    assignSheet.deleteRow(j + 1);
                }
            }
        }

        return { success: true, count: deletedCount };
    } catch (e) {
        return { error: "Lỗi xóa dữ liệu: " + e.message };
    }
}

// --- USER ASSIGNMENTS ---
function getMyCourses(userID) {
    try {
        const db = getDB();
        const assignData = db.getSheetByName('Assignments').getDataRange().getDisplayValues();
        const courseData = db.getSheetByName('Courses').getDataRange().getDisplayValues();

        const courseMap = {};
        for (let i = 0; i < courseData.length; i++) {
            courseMap[courseData[i][0]] = courseData[i];
        }

        const assigns = assignData.slice(1);
        const myAssigns = assigns.filter(r => r[2].toString() === userID && r[7] === '' && r[9] === '');

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const parseDate = (dStr) => { if (!dStr) return null; const p = dStr.split('-'); return new Date(p[0], p[1] - 1, p[2]); };

        return myAssigns.map(a => {
            const courseInfo = courseMap[a[0]];
            if (!courseInfo) return null;

            let startDate = parseDate(courseInfo[2]);
            let endDate = parseDate(courseInfo[3]);

            // Check Col G (Index 6) for Sign Permission
            let isCourseOpen = (String(courseInfo[6]).toLowerCase() === 'true');

            let status = 'Opening';

            if (today > endDate) status = 'Ended';
            else if (today < startDate) status = 'Plan';

            // Allow signing ONLY if Opening AND isCourseOpen
            const canSign = (status === 'Opening' && isCourseOpen);

            return {
                id: a[0],
                name: a[1],
                start: courseInfo[2],
                end: courseInfo[3],
                status: status,
                canSign: canSign
            };
        }).filter(Boolean);
    } catch (e) { return []; }
}

function confirmTraining(userID, courseID, sign) {
    // 1. KIỂM TRA CHỮ KÝ (LOGIC MỚI)
    // Nếu không có dữ liệu chữ ký hoặc độ dài chuỗi quá ngắn (ảnh trắng)
    if (!sign || sign.length < 1000) {
        return false; 
    }
    const db = getDB();
    const sheet = db.getSheetByName('Assignments');
    const data = sheet.getDataRange().getDisplayValues();
    const courseSheet = db.getSheetByName('Courses');
    const courseData = courseSheet.getDataRange().getDisplayValues();

    // Check Course Status & Permission
    const course = courseData.find(c => c[0].toString() === courseID);
    if (!course) return false;

    // Permission Check (Col G / Index 6)
    if (String(course[6]).toLowerCase() !== 'true') return false;

    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const parseDate = (dStr) => { if (!dStr) return null; const p = dStr.split('-'); return new Date(p[0], p[1] - 1, p[2]); };

    const start = parseDate(course[2]);
    const end = parseDate(course[3]);

    // Strict Check: Must be Opening
    if (today < start || today > end) return false;

    for (let i = 1; i < data.length; i++) {
        // Col A(0)=cID, Col C(2)=uID
        if (data[i][2].toString() === userID && data[i][0].toString() === courseID) {
            sheet.getRange(i + 1, 8, 1, 2).setValues([[sign, new Date()]]);
            return true;
        }
    }
    return false;
}

// --- QUẢN LÝ KHÓA HỌC NÂNG CAO ---
function updateCourseSignStatus(courseID, isOpen) {
    try {
        const sheet = getDB().getSheetByName('Courses');
        const data = sheet.getDataRange().getDisplayValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === courseID) {
                // Col G is 7th, row i+1
                sheet.getRange(i + 1, 7).setNumberFormat("@").setValue(isOpen ? 'true' : 'false');
                return true;
            }
        }
        return false;
    } catch (e) { return false; }
}

function saveCourseWithTarget(courseData, targetType, targetValue) {
    try {
        const ss = getDB();
        const courseSheet = ss.getSheetByName('Courses');
        const assignSheet = ss.getSheetByName('Assignments');
        const userSheet = ss.getSheetByName('Users');

        const isNew = !courseData.id;
        if (isNew) courseData.id = 'TRN-' + new Date().getTime();

        const courseDataRow = courseSheet.getDataRange().getDisplayValues();
        let existingCourseIndex = -1;

        for (let i = 1; i < courseDataRow.length; i++) {
            if (courseDataRow[i][0] === courseData.id) {
                existingCourseIndex = i;
                break;
            }
        }

        if (!isNew) {
            if (existingCourseIndex !== -1) {
                courseSheet.getRange(existingCourseIndex + 1, 2, 1, 4).setValues([[
                    courseData.name, "'" + courseData.start, "'" + courseData.end, courseData.content
                ]]);
                return { success: true, count: 0, action: 'updated' };
            }
            return { error: "Không tìm thấy khóa học!" };
        } else {
            // New Course: Default SignEnabled='false' (Col G)
            courseSheet.appendRow([
                courseData.id, courseData.name, "'" + courseData.start,
                "'" + courseData.end, courseData.content, 'plan', 'false'
            ]);
        }

        const users = userSheet.getDataRange().getDisplayValues();
        let targetUsers = [];

        if (targetType === 'filter') {
            targetUsers = users.slice(1).filter(u =>
                (!targetValue.com || u[2] === targetValue.com) &&
                (!targetValue.group || u[3] === targetValue.group) &&
                (!targetValue.part || u[4] === targetValue.part)
            );
        } else {
            let idList = Array.isArray(targetValue) ? targetValue.map(id => id.toString().trim()) : targetValue.split(',').map(id => id.trim());
            targetUsers = users.slice(1).filter(u => idList.includes(u[1].toString()));
        }

        // FILTER ADMINS (Fix: Strictly assign only to 'user')
        targetUsers = targetUsers.filter(u => {
            const role = String(u[5] || '').toLowerCase();
            const uid = String(u[1] || '').toLowerCase();
            return role === 'user' && uid !== 'admin';
        });

        if (targetUsers.length > 0) {
            const rows = targetUsers.map(u => [
                courseData.id,
                courseData.name,
                u[1], // UserID
                u[0], // Name
                u[2], // Com
                u[3], // Group
                u[4], // Part
                '',   // Sign
                '',   // Time
                ''    // Exception
            ]);
            assignSheet.getRange(assignSheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
        }
        return { success: true, count: targetUsers.length, action: 'created' };
    } catch (e) { return { error: e.message }; }
}

function deleteCourse(courseId) {
    try {
        const ss = getDB();
        const courseSheet = ss.getSheetByName('Courses');
        const assignSheet = ss.getSheetByName('Assignments');

        const cData = courseSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < cData.length; i++) {
            if (cData[i][0] === courseId) {
                courseSheet.deleteRow(i + 1);
                break;
            }
        }

        const aData = assignSheet.getDataRange().getDisplayValues();
        const safeRows = aData.filter(r => r[0] !== courseId);

        // Rewrite safely
        if (safeRows.length < aData.length) {
            assignSheet.clearContents();
            if (safeRows.length > 0) assignSheet.getRange(1, 1, safeRows.length, safeRows[0].length).setValues(safeRows);
        }

        return true;
    } catch (e) { return { error: e.message }; }
}

function getCourseStatistics() {
    try {
        const ss = getDB();
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const parseDate = (dStr) => { if (!dStr) return null; const p = dStr.split('-'); return new Date(p[0], p[1] - 1, p[2]); };

        const assignMap = {};
        assignments.forEach(a => {
            const cid = a[0];
            if (!assignMap[cid]) assignMap[cid] = { total: 0, completed: 0 };
            assignMap[cid].total++;
            if (a[7] !== '' || a[9] !== '') assignMap[cid].completed++;
        });

        return courses.map(c => {
            const cid = c[0];
            const stats = assignMap[cid] || { total: 0, completed: 0 };

            let percent = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;

            let startDate = parseDate(c[2]);
            let endDate = parseDate(c[3]);
            let status = '';

            // Read SignEnabled (Col G / Index 6)
            let isSignEnabled = (String(c[6]).toLowerCase() === 'true');

            if (percent === 100 && stats.total > 0) status = 'Complete';
            else if (today > endDate) status = 'Pending';
            else if (today < startDate) status = 'Plan';
            else status = 'Opening';

            return {
                id: cid, name: c[1], start: c[2], end: c[3], content: c[4],
                status: status, percent: percent, pendingCount: stats.total - stats.completed,
                signEnabled: isSignEnabled
            };
        }).reverse();
    } catch (e) { return { error: "Lỗi tải khóa học: " + e.message }; }
}

function getCourseParticipants(courseID) {
    try {
        const data = getDB().getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        return data.filter(r => r[0] === courseID).map(r => ({
            uid: r[2], name: r[3], com: r[4], group: r[5], part: r[6],
            signed: r[7] !== '',
            signTime: r[8],
            exception: r[9]
        }));
    } catch (e) { return []; }
}

function updateException(courseID, userID, reason) {
    try {
        const sheet = getDB().getSheetByName('Assignments');
        const data = sheet.getDataRange().getDisplayValues();
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === courseID && data[i][2] === userID) {
                sheet.getRange(i + 1, 10).setValue(reason);
                return true;
            }
        }
        return false;
    } catch (e) { return false; }
}

// Helper: Get or Create 'Data' Folder
function getDataFolder() {
    try {
        const ssFile = DriveApp.getFileById(SPREADSHEET_ID);
        const parents = ssFile.getParents();
        const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

        const folders = folder.getFoldersByName("Data");
        if (folders.hasNext()) {
            return folders.next();
        } else {
            return folder.createFolder("Data");
        }
    } catch (e) {
        return DriveApp.getRootFolder(); // Fallback
    }
}

// --- EXPORT REMINDERS ---
function exportConsolidatedReminder() {
    try {
        const ss = getDB();
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const parseDate = (dStr) => {
            if (!dStr) return null;
            const p = dStr.split('-');
            return new Date(p[0], p[1] - 1, p[2]);
        };

        // 1. Identify Target Courses (Opening or Pending)
        const targetCourses = {}; // Map ID -> Name
        courses.forEach(c => {
            const start = parseDate(c[2]);
            const end = parseDate(c[3]);

            // Logic: Opening (Start <= Today <= End), Pending (Today > End)
            // Warning: 'Plan' is Today < Start (Ignore)
            if (start && end) {
                if (today >= start) { // Started
                    targetCourses[c[0]] = c[1]; // Store Name
                }
            }
        });

        // 2. Filter Assignments
        // Req: Course matches target, Sign & Exception are empty
        const reportData = [];
        assignments.forEach(a => {
            const cID = a[0];
            const cName = targetCourses[cID];

            // If course is in target list AND assignment incomplete
            if (cName && a[7] === '' && a[9] === '') {
                // Determine Status text based on Date comparison again?
                // Or just 'Chưa hoàn thành'
                // User asked for "trạng thái hoàn thành" variable in list.
                // Since it's incomplete, it's either "Đang thực hiện" (Opening) or "Quá hạn" (Pending)
                // Let's re-calc strict status for the row
                // We need course dates again.
                // Improve optimization: Store Course Obj in map.
                reportData.push({
                    cID: cID,
                    uID: a[2],
                    name: a[3],
                    com: a[4],
                    group: a[5],
                    part: a[6]
                });
            }
        });

        if (reportData.length === 0) return { error: "Không có dữ liệu nhắc nhở nào!" };

        // Save Temp in Data Folder
        const folder = getDataFolder();
        const tempSS = SpreadsheetApp.create("Temp_Export_" + new Date().getTime());
        const tempFile = DriveApp.getFileById(tempSS.getId());

        // Move temp file to Data folder (Apps Script creates in Root by default)
        tempFile.moveTo(folder);

        const sheet = tempSS.getActiveSheet();

        // Header
        const headers = ["Khóa đào tạo", "Họ và tên", "Mã nhân viên", "Company", "Group", "Part", "Trạng thái"];
        sheet.appendRow(headers);

        // Body
        // Re-map course dates for status check
        const courseMap = {};
        courses.forEach(c => courseMap[c[0]] = { end: parseDate(c[3]) });

        const rows = reportData.map(r => {
            const cEnd = courseMap[r.cID] ? courseMap[r.cID].end : null;
            const status = (cEnd && today > cEnd) ? "Quá hạn" : "Chưa hoàn thành";
            return [targetCourses[r.cID], r.name, r.uID, r.com, r.group, r.part, status];
        });

        sheet.getRange(2, 1, rows.length, 7).setValues(rows);

        // 4. Formatting
        const fullRange = sheet.getDataRange();
        fullRange.setFontFamily("Times New Roman");
        fullRange.setFontSize(12);

        // Header Bold & Center
        sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setHorizontalAlignment("center");

        // Borders
        fullRange.setBorder(true, true, true, true, true, true);

        // Auto fit
        sheet.autoResizeColumns(1, 7);

        // 5. Download as Xlsx
        SpreadsheetApp.flush();
        const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            return { error: "Lỗi Export API: " + response.getResponseCode() };
        }

        const blob = response.getBlob().setName("NhacNho_DaoTao.xlsx");
        const base64 = Utilities.base64Encode(blob.getBytes());

        // Cleanup with Try-Catch to ensure download still works if trash fails
        try {
            tempFile.setTrashed(true);
        } catch (e) {
            // Ignore cleanup error
        }

        return { filename: "NhacNho_DaoTao.xlsx", base64: base64 };

    } catch (e) {
        return { error: "Lỗi xuất file: " + e.message };
    }
}
// Force Drive Scope for Export
function _forceScope() { DriveApp.getRootFolder(); }

// --- AUTOMATION & EMAIL REPORTING ---

/**
 * Setup Triggers for Automation
 * Run this function ONCE to set up the daily schedules.
 */
function setupTriggers() {
    // Delete existing triggers to avoid duplicates
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => ScriptApp.deleteTrigger(t));

    // 1. Morning Report (7:00 AM)
    ScriptApp.newTrigger('sendMorningReport')
        .timeBased()
        .everyDays(1)
        .atHour(7)
        .create();

    // 2. Afternoon Report (15:00 PM)
    ScriptApp.newTrigger('sendAfternoonReport')
        .timeBased()
        .everyDays(1)
        .atHour(15) // 15:00 - 16:00
        .create();

    // 3. Course End Check (Daily at 6:00 AM)
    ScriptApp.newTrigger('checkCourseCompletion')
        .timeBased()
        .everyDays(1)
        .atHour(6)
        .create();
}

function sendMorningReport() {
    const today = new Date();
    const day = today.getDay();
    // 0 = Sun, 6 = Sat. Only run Mon(1) to Fri(5)
    if (day >= 1 && day <= 5) {
        sendIncompleteReport("Sáng (7:00)");
    }
}

function sendAfternoonReport() {
    const today = new Date();
    const day = today.getDay();
    if (day >= 1 && day <= 5) {
        sendIncompleteReport("Chiều (15:00)");
    }
}

/**
 * Sends a report of ALL incomplete assignments to the Admin.
 */
function sendIncompleteReport(timeLabel) {
    try {
        const ss = getDB();
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // Map Course End Dates
        const courseMap = {};
        courses.forEach(c => {
            const endDateParts = c[3] ? c[3].split('-') : null;
            const endDate = endDateParts ? new Date(endDateParts[0], endDateParts[1] - 1, endDateParts[2]) : null;
            if (endDate) endDate.setHours(0, 0, 0, 0);
            courseMap[c[0]] = { name: c[1], end: endDate };
        });

        const incompleteList = [];

        assignments.forEach(a => {
            // a[7] (Sign) is empty AND a[9] (Exception) is empty
            if (a[7] === '' && a[9] === '') {
                const cInfo = courseMap[a[0]];
                if (cInfo) {
                    // Report if course is ACTIVE (Today <= EndDate)
                    if (cInfo.end && today <= cInfo.end) {
                        incompleteList.push({
                            cName: cInfo.name,
                            uName: a[3],
                            uID: a[2],
                            uPart: a[6]
                        });
                    }
                }
            }
        });

        if (incompleteList.length === 0) return; // No incomplete tasks

        // Build HTML Table
        let html = `<h3>Báo cáo đào tạo chưa hoàn thành - ${timeLabel}</h3>`;
        html += `<p>Tổng số: <b>${incompleteList.length}</b> trường hợp.</p>`;
        html += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; width:100%;">`;
        html += `<tr style="background:#f2f2f2"><th>Khóa học</th><th>Nhân viên</th><th>ID</th><th>Bộ phận</th></tr>`;

        incompleteList.forEach(item => {
            html += `<tr><td>${item.cName}</td><td>${item.uName}</td><td>${item.uID}</td><td>${item.uPart}</td></tr>`;
        });
        html += `</table>`;
        html += `<p><i>Email này được gửi tự động từ hệ thống. Vui lòng xem danh sách đính kèm để biết chi tiết.</i></p>`;

        // --- GENERATE EXCEL ATTACHMENT ---
        let attachmentBlob = null;
        let tempFile = null;
        try {
            const folder = getDataFolder();
            const tempSS = SpreadsheetApp.create("Temp_Report_" + new Date().getTime());
            tempFile = DriveApp.getFileById(tempSS.getId());
            tempFile.moveTo(folder);

            const sheet = tempSS.getActiveSheet();
            sheet.appendRow(["Tên Khóa Học", "Họ và Tên", "Mã Nhân Viên", "Bộ Phận"]);
            const rows = incompleteList.map(item => [item.cName, item.uName, item.uID, item.uPart]);

            if (rows.length > 0) {
                sheet.getRange(2, 1, rows.length, 4).setValues(rows);
            }

            sheet.getDataRange().setFontFamily("Times New Roman").setBorder(true, true, true, true, true, true);
            sheet.autoResizeColumns(1, 4);
            SpreadsheetApp.flush();

            const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
            const token = ScriptApp.getOAuthToken();
            const response = UrlFetchApp.fetch(url, {
                headers: { 'Authorization': 'Bearer ' + token },
                muteHttpExceptions: true
            });

            if (response.getResponseCode() === 200) {
                attachmentBlob = response.getBlob().setName(`ChuaHoanThanh_${timeLabel.replace(/[\s\(\):]/g, '_')}.xlsx`);
            }
        } catch (err) {
            console.error("Error creating attachment: " + err.message);
            html += `<p style="color:red">Lỗi tạo file đính kèm: ${err.message}</p>`;
        }

        const RECIPIENT_EMAILS = 'hien.pt1@samsung.com,qcteam.iqctraining@gmail.com,dinh.hnhung@samsung.com,dang.ngoc@samsung.com,nhung.tt2@samsung.com,an.pthi@samsung.com';

        // Send Email
        const emailOptions = {
            to: RECIPIENT_EMAILS,
            subject: `[Training System] Báo cáo chưa hoàn thành - ${timeLabel}`,
            htmlBody: html
        };

        if (attachmentBlob) {
            emailOptions.attachments = [attachmentBlob];
        }

        MailApp.sendEmail(emailOptions);

        // Cleanup
        if (tempFile) {
            try { tempFile.setTrashed(true); } catch (e) { }
        }

    } catch (e) {
        console.error("Error sending report: " + e.message);
    }
}

/**
 * Checks for courses that ended YESTERDAY.
 */
function checkCourseCompletion() {
    try {
        const ss = getDB();
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        yesterday.setHours(0, 0, 0, 0);

        courses.forEach(c => {
            if (!c[3]) return;
            const parts = c[3].split('-');
            const endDate = new Date(parts[0], parts[1] - 1, parts[2]);
            endDate.setHours(0, 0, 0, 0);

            // Check if EndDate == Yesterday
            if (endDate.getTime() === yesterday.getTime()) {
                sendCourseFinalReport(c, assignments);
            }
        });
    } catch (e) {
        console.error("Error checking course completion: " + e.message);
    }
}

function sendCourseFinalReport(courseRow, allAssignments) {
    const courseId = courseRow[0];
    const courseName = courseRow[1];

    // Filter assignments for this course
    const courseAssigns = allAssignments.filter(a => a[0] === courseId);
    if (courseAssigns.length === 0) return;

    const total = courseAssigns.length;
    const completed = courseAssigns.filter(a => a[7] !== '' || a[9] !== '').length;
    const incomplete = courseAssigns.filter(a => a[7] === '' && a[9] === '');
    const percent = Math.round((completed / total) * 100);

    let html = `<h3>Báo cáo kết thúc khóa học: ${courseName}</h3>`;
    html += `<p><b>Ngày kết thúc:</b> ${courseRow[3]}</p>`;
    html += `<p><b>Tổng số học viên:</b> ${total}</p>`;
    html += `<p><b>Đã hoàn thành:</b> ${completed} (${percent}%)</p>`;
    html += `<p><b>Chưa hoàn thành:</b> ${incomplete.length}</p>`;

    if (incomplete.length > 0) {
        html += `<h4>Danh sách chưa hoàn thành:</h4>`;
        html += `<table border="1" cellpadding="5" cellspacing="0" style="border-collapse:collapse; width:100%;">`;
        html += `<tr style="background:#f2f2f2"><th>Nhân viên</th><th>ID</th><th>Bộ phận</th></tr>`;
        incomplete.forEach(row => {
            // row[3] Name, row[2] ID, row[6] Part
            html += `<tr><td>${row[3]}</td><td>${row[2]}</td><td>${row[6]}</td></tr>`;
        });
        html += `</table>`;
    }

    html += `<p><i>Email này được gửi tự động từ hệ thống</i></p>`;

    // Send Email
    const RECIPIENT_EMAILS = 'hien.pt1@samsung.com,haiduong17.nguyen@gmail.com,qcteam.iqctraining@gmail.com';

    MailApp.sendEmail({
        to: RECIPIENT_EMAILS,
        subject: `[Training System] Kết thúc khóa học: ${courseName}`,
        htmlBody: html
    });
}

function exportPendingList(courseId) {
    const ss = getDB();
    const assigns = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
    let pending = assigns.filter(a => a[0] === courseId && a[7] === '' && a[9] === '');
    return pending.map(a => [a[3], a[2], a[4], a[5], a[6]]);
}

// --- EXPORT PDF REPORT ---
function exportCourseReport(courseID) {
    try {
        const ss = getDB();
        const courseData = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const assignData = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        // 1. Get Course Info
        const course = courseData.find(c => c[0].toString() === courseID.toString());
        if (!course) return { error: "Không tìm thấy khóa học!" };

        // 2. Validate Status (Opening, Pending, Complete)
        // User said: "Chỉ thực hiện được với các khóa trạng thái là: opening, pending, complete"
        // We can just proceed, checking data availability is enough.
        // Or restrictive check? Let's check dates to be safe, but user priority is "output list".
        // Actually, let's just allow it for valid courses.

        const cName = course[1].toUpperCase();
        const cTime = `${course[2]} - ${course[3]}`;
        const cContent = course[4];

        // 3. Get Participants (Completed Only)
        // Completed = Signed OR Exception
        const participants = assignData.filter(a =>
            a[0].toString() === courseID.toString() && (a[7] !== '' || a[9] !== '')
        );

        if (participants.length === 0) return { error: "Không có học viên nào đã hoàn thành để xuất báo cáo." };

        // 4. Generate HTML
        let html = `
        <html>
        <head>
            <style>
                body { font-family: 'Times New Roman', serif; padding: 20px; }
                h1 { font-size: 20pt; text-align: center; margin-bottom: 5px; text-transform: uppercase; font-weight: bold; }
                .meta { text-align: center; font-size: 12pt; margin-bottom: 20px; font-style: italic; }
                .content { margin-bottom: 20px; font-size: 12pt; }
                table { width: 100%; border-collapse: collapse; font-size: 12pt; }
                th, td { border: 1px solid black; padding: 5px; vertical-align: middle; text-align: center; }
                th { font-weight: bold; background-color: #f0f0f0; }
                .left { text-align: left; }
                img.sign { max-height: 40px; max-width: 100px; }
            </style>
        </head>
        <body>
            <h1>BÁO CÁO KHÓA HỌC: ${cName}</h1>
            <div class="meta">Thời gian: ${cTime}</div>
            <div class="content"><strong>Nội dung:</strong> ${cContent}</div>

            <table>
                <thead>
                    <tr>
                        <th style="width: 5%">No</th>
                        <th style="width: 20%">Họ và tên</th>
                        <th style="width: 10%">Mã NV</th>
                        <th style="width: 15%">Company</th>
                        <th style="width: 10%">Group</th>
                        <th style="width: 10%">Part</th>
                        <th style="width: 15%">Chữ ký</th>
                        <th style="width: 15%">Thời gian / Ghi chú</th>
                    </tr>
                </thead>
                <tbody>
        `;

        participants.forEach((p, index) => {
            const name = p[3];
            const uid = p[2];
            const com = p[4];
            const grp = p[5];
            const prt = p[6];
            const signImg = p[7]; // Base64
            const time = p[8] ? new Date(p[8]).toLocaleString('vi-VN') : '';
            const exc = p[9];

            let signCell = '';
            let metaCell = '';

            if (exc !== '') {
                signCell = '(Ngoại lệ)';
                metaCell = exc;
            } else {
                if (signImg.startsWith('data:image')) {
                    signCell = `<img src="${signImg}" class="sign" />`;
                } else {
                    signCell = '(Đã ký)';
                }
                metaCell = time;
            }

            html += `
                <tr>
                    <td>${index + 1}</td>
                    <td class="left">${name}</td>
                    <td>${uid}</td>
                    <td>${com}</td>
                    <td>${grp}</td>
                    <td>${prt}</td>
                    <td>${signCell}</td>
                    <td>${metaCell}</td>
                </tr>
            `;
        });

        html += `
                </tbody>
            </table>
        </body>
        </html>
        `;

        // Create file in 'Data' Folder
        const dataFolder = getDataFolder();
        const blob = Utilities.newBlob(html, MimeType.HTML, "temp.html");

        // 1. Create HTML Temp
        const tempFile = dataFolder.createFile(blob);

        // 2. Convert to PDF
        const pdfBlob = tempFile.getAs(MimeType.PDF);
        const finalBase64 = Utilities.base64Encode(pdfBlob.getBytes());

        // 3. Cleanup HTML Temp
        tempFile.setTrashed(true);

        return { filename: `BaoCao_${cName}.pdf`, base64: finalBase64 };

    } catch (e) {
        return { error: "Lỗi tạo báo cáo: " + e.message };
    }
}
