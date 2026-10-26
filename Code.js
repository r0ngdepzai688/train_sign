/**
 * IQC NEXUS - HỆ THỐNG QUẢN LÝ VÀ KÝ XÁC NHẬN ĐÀO TẠO
 * Phiên bản: 2.1 (Ultra Low Latency & High Concurrency Architecture)
 * Tác giả: Nguyễn Hải Đường (hai.duong)
 * Slogan: Silence the ads. Let quality make the noise.
 */

// --- CẤU HÌNH HỆ THỐNG ---
const ADMIN_PASS_DEFAULT = "@nhd16488";
const USER_DEFAULT_PASS_DEFAULT = "@iqc6368";
const SPREADSHEET_ID_DEFAULT = "15K3FetZvSMPRKLNp9I7kyPLCeAIpreWNeeG7uerW3wk";
const RECIPIENT_EMAILS_DEFAULT = 'hien.pt1@samsung.com,qcteam.iqctraining@gmail.com,dinh.hnhung@samsung.com,dang.ngoc@samsung.com,nhung.tt2@samsung.com,an.pthi@samsung.com,haiduong17.nguyen@gmail.com';

function getProp(key, fallback) {
    try {
        const val = PropertiesService.getScriptProperties().getProperty(key);
        return (val !== null && val !== undefined && val !== '') ? val : fallback;
    } catch (e) {
        return fallback;
    }
}

function getDB() {
    const sheetId = getProp('SPREADSHEET_ID', SPREADSHEET_ID_DEFAULT);
    return SpreadsheetApp.openById(sheetId);
}

function doGet() {
    return HtmlService.createTemplateFromFile('Index')
        .evaluate()
        .setTitle('IQC Nexus | Hệ Thống Đào Tạo & Ký Số')
        .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0, viewport-fit=cover')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// --- BẢO MẬT PHÂN QUYỀN & TOKEN CACHE ---

function verifyAdminToken(token) {
    if (!token) return false;
    try {
        const cached = CacheService.getScriptCache().get('AUTH_' + token);
        return cached === 'admin';
    } catch (e) {
        return false;
    }
}

/**
 * Đăng nhập siêu tốc (Zero-Roundtrip):
 * Trả về luôn dữ liệu khởi tạo (Dashboard cho Admin hoặc Danh sách khóa học cho User)
 * Giúp giao diện hiển thị ngay lập tức, giảm 50% thời gian chờ tải.
 */
function checkLogin(user, pass) {
    try {
        const adminPass = getProp('ADMIN_PASS', ADMIN_PASS_DEFAULT);
        const userDefaultPass = getProp('USER_DEFAULT_PASS', USER_DEFAULT_PASS_DEFAULT);
        const uTarget = String(user).trim();

        // 1. Quản trị viên
        if (uTarget === 'admin' && pass === adminPass) {
            const token = Utilities.getUuid() + '-' + Date.now();
            try {
                CacheService.getScriptCache().put('AUTH_' + token, 'admin', 21600); // 6 giờ
            } catch (e) { }

            // Tải sẵn dữ liệu dashboard admin để render 0ms ngay sau khi đăng nhập
            const dashboardData = getAdminDashboardData(token);
            return {
                role: 'admin',
                user: 'Quản trị viên',
                userID: 'admin',
                token: token,
                initialDashboard: dashboardData
            };
        }

        // 2. Nhân viên IQC
        const ss = getDB();
        const userSheet = ss.getSheetByName('Users');
        if (!userSheet) return { error: "Không tìm thấy bảng Users!" };

        const userData = userSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < userData.length; i++) {
            let uID = String(userData[i][1] || '').trim();
            if (uID === uTarget && pass === userDefaultPass) {
                if (userData[i][6] === 'Inactive') return { error: "Tài khoản của bạn đã bị khóa!" };

                // Tải sẵn danh sách khóa học của nhân viên này để loại bỏ round-trip thứ 2
                const myCourses = getMyCourses(uID);
                return {
                    role: 'user',
                    user: userData[i][0],
                    userID: uID,
                    courses: myCourses
                };
            }
        }
        return { error: "Mã nhân viên hoặc mật khẩu không chính xác!" };
    } catch (e) {
        return { error: "Lỗi máy chủ: " + e.message };
    }
}

// --- UTILITIES: DATE PARSER & CACHE MANAGEMENT ---

function parseDateFlexible(dVal) {
    if (!dVal) return null;
    if (dVal instanceof Date) {
        const d = new Date(dVal.getTime());
        d.setHours(0, 0, 0, 0);
        return d;
    }
    const str = String(dVal).trim();
    if (!str) return null;

    if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(str)) {
        const p = str.split('-');
        return new Date(parseInt(p[0], 10), parseInt(p[1], 10) - 1, parseInt(p[2], 10));
    }
    if (/^\d{1,2}\/\d{1,2}\/\d{4}$/.test(str)) {
        const p = str.split('/');
        return new Date(parseInt(p[2], 10), parseInt(p[1], 10) - 1, parseInt(p[0], 10));
    }
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
        d.setHours(0, 0, 0, 0);
        return d;
    }
    return null;
}

function clearDashboardCache() {
    try {
        const cache = CacheService.getScriptCache();
        cache.remove('admin_filters');
        cache.remove('available_courses');
        cache.remove('courses_map_cache');
    } catch (e) { }
}

// --- DRIVE STORAGE ENGINE CHO CHỮ KÝ ---

function getDataFolder() {
    try {
        const sheetId = getProp('SPREADSHEET_ID', SPREADSHEET_ID_DEFAULT);
        const ssFile = DriveApp.getFileById(sheetId);
        const parents = ssFile.getParents();
        const folder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

        const folders = folder.getFoldersByName("Data");
        if (folders.hasNext()) {
            return folders.next();
        } else {
            return folder.createFolder("Data");
        }
    } catch (e) {
        return DriveApp.getRootFolder();
    }
}

function getSignaturesFolder() {
    try {
        const dataFolder = getDataFolder();
        const folders = dataFolder.getFoldersByName("Signatures");
        if (folders.hasNext()) {
            return folders.next();
        } else {
            return dataFolder.createFolder("Signatures");
        }
    } catch (e) {
        return getDataFolder();
    }
}

function saveSignatureToDrive(courseID, userID, signatureDataUrl) {
    if (!signatureDataUrl || !signatureDataUrl.startsWith('data:image')) {
        return signatureDataUrl || '';
    }
    try {
        const folder = getSignaturesFolder();
        const base64Content = signatureDataUrl.replace(/^data:image\/\w+;base64,/, '');
        const decodedBytes = Utilities.base64Decode(base64Content);
        const fileName = `SIG_${courseID}_${userID}_${Date.now()}.png`;
        const blob = Utilities.newBlob(decodedBytes, MimeType.PNG, fileName);
        const file = folder.createFile(blob);
        return "drive:" + file.getId();
    } catch (err) {
        console.error("Lỗi lưu Drive: " + err.message);
        return signatureDataUrl; // Fallback giữ an toàn
    }
}

function resolveSignatureImageBase64(signVal) {
    if (!signVal) return '';
    if (signVal.startsWith('drive:')) {
        const fileId = signVal.substring(6).trim();
        try {
            const file = DriveApp.getFileById(fileId);
            return 'data:image/png;base64,' + Utilities.base64Encode(file.getBlob().getBytes());
        } catch (e) {
            return '';
        }
    }
    return signVal;
}

// --- DATA ACCESS & DASHBOARD APIS ---

function getAdminUsers(adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const ss = getDB();
        const sheet = ss.getSheetByName('Users');
        if (!sheet) return [];
        return sheet.getDataRange().getDisplayValues().slice(1);
    } catch (e) {
        return { error: e.toString() };
    }
}

function getAdminCourses(adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    return getCourseStatistics();
}

function getAdminFilters(adminToken) {
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('admin_filters');
        if (cached) return JSON.parse(cached);

        const users = getDB().getSheetByName('Users').getDataRange().getValues().slice(1);
        const getUnique = (idx) => [...new Set(users.map(u => String(u[idx] || '').trim()).filter(Boolean))].sort();

        const result = {
            com: getUnique(2),
            group: getUnique(3),
            part: getUnique(4)
        };
        cache.put('admin_filters', JSON.stringify(result), 900); // 15 phút
        return result;
    } catch (e) {
        return { com: [], group: [], part: [] };
    }
}

function getAdminDashboardData(adminToken) {
    if (!verifyAdminToken(adminToken)) {
        return { error: "Phiên làm việc quản trị đã hết hạn hoặc không hợp lệ. Vui lòng đăng nhập lại!" };
    }
    const users = getAdminUsers(adminToken);
    const newRegs = Array.isArray(users) ? users.filter(u => u[6] === 'New-Reg') : [];

    return {
        courses: getCourseStatistics(),
        filters: getAdminFilters(adminToken),
        newRegisters: newRegs,
        availableCourses: getAvailableCourses(),
        allUsers: users
    };
}

function getCourseDetails(courseId) {
    try {
        const db = getDB();
        const sheet = db.getSheetByName('Courses');
        if (!sheet) return '';
        const data = sheet.getDataRange().getValues();
        const course = data.find(r => String(r[0]).trim() === String(courseId).trim());
        return course ? course[4] : '';
    } catch (e) {
        return '';
    }
}

function getAvailableCourses() {
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('available_courses');
        if (cached) return JSON.parse(cached);

        const ss = getDB();
        const courseSheet = ss.getSheetByName('Courses');
        if (!courseSheet) return [];
        const courseData = courseSheet.getDataRange().getDisplayValues().slice(1);
        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const list = courseData.map(c => {
            const start = parseDateFlexible(c[2]);
            const end = parseDateFlexible(c[3]);
            let status = 'Opening';
            if (end && today > end) status = 'Pending';
            else if (start && today < start) status = 'Plan';

            return { id: c[0], name: c[1], status: status };
        });

        cache.put('available_courses', JSON.stringify(list), 600); // 10 phút
        return list;
    } catch (e) {
        return [];
    }
}

// --- USER MANAGEMENT (BATCH CRUD) ---

function saveUser(u, isEdit, courseIds, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const rowData = [u.name, u.id, u.com, u.group, u.part, 'user', u.status || 'Active'];

        if (isEdit) {
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][1]).trim() === String(u.id).trim()) {
                    sheet.getRange(i + 1, 2).setNumberFormat("@");
                    sheet.getRange(i + 1, 1, 1, 7).setValues([rowData]);
                    return { success: true, user: rowData, action: 'update' };
                }
            }
            return { error: "Không tìm thấy người dùng để cập nhật!" };
        } else {
            for (let i = 1; i < data.length; i++) {
                if (String(data[i][1]).trim() === String(u.id).trim()) {
                    return { error: "Mã nhân viên đã tồn tại!" };
                }
            }
            const newRow = sheet.getLastRow() + 1;
            sheet.getRange(newRow, 2).setNumberFormat("@");
            sheet.getRange(newRow, 1, 1, 7).setValues([rowData]);

            if (courseIds && courseIds.length > 0) {
                assignCoursesToUsers([u.id], courseIds, adminToken);
            }

            return { success: true, user: rowData, action: 'create' };
        }
    } catch (e) {
        return { error: "Lỗi máy chủ: " + e.message };
    }
}

function deleteUser(userID, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const ss = getDB();
        const sheet = ss.getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const targetId = String(userID).trim();

        if (targetId.toLowerCase() === 'admin') {
            return { error: "Không thể xóa tài khoản Quản trị viên!" };
        }

        let userRowIndex = -1;
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][1]).trim() === targetId) {
                userRowIndex = i + 1;
                break;
            }
        }

        if (userRowIndex === -1) return { error: "Không tìm thấy nhân viên!" };

        const assignSheet = ss.getSheetByName('Assignments');
        if (assignSheet) {
            const aData = assignSheet.getDataRange().getDisplayValues();
            if (aData.length > 1) {
                const header = aData[0];
                const safeAssigns = [header];
                for (let j = 1; j < aData.length; j++) {
                    const isTarget = String(aData[j][2]).trim() === targetId;
                    const isIncomplete = aData[j][7] === '' && aData[j][9] === '';
                    if (!(isTarget && isIncomplete)) {
                        safeAssigns.push(aData[j]);
                    }
                }
                if (safeAssigns.length !== aData.length) {
                    assignSheet.getRange(1, 1, aData.length, aData[0].length).clearContent();
                    assignSheet.getRange(1, 1, safeAssigns.length, safeAssigns[0].length).setValues(safeAssigns);
                }
            }
        }

        sheet.deleteRow(userRowIndex);
        return { success: true, id: userID };
    } catch (e) {
        return { error: e.message };
    }
}

function registerUser(u) {
    try {
        clearDashboardCache();
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const targetId = String(u.id || '').trim();

        if (!targetId || !u.name) return { error: "Vui lòng nhập đầy đủ họ tên và mã nhân viên!" };

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][1]).trim() === targetId) {
                return { error: "Mã nhân viên đã tồn tại trên hệ thống!" };
            }
        }

        const nextRow = sheet.getLastRow() + 1;
        const newRowData = [u.name, targetId, u.com || '', u.group || '', u.part || '', 'user', 'New-Reg'];

        sheet.getRange(nextRow, 2).setNumberFormat("@");
        sheet.getRange(nextRow, 1, 1, 7).setValues([newRowData]);
        return { success: true };
    } catch (e) {
        return { error: "Lỗi đăng ký: " + e.message };
    }
}

function assignCoursesToUsers(userIds, courseIds, adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    if (!userIds || userIds.length === 0 || !courseIds || courseIds.length === 0) return { count: 0 };
    try {
        const ss = getDB();
        const assignSheet = ss.getSheetByName('Assignments');
        const userSheet = ss.getSheetByName('Users');
        const courseSheet = ss.getSheetByName('Courses');

        const allUsers = userSheet.getDataRange().getDisplayValues();
        const userMap = {};
        allUsers.slice(1).forEach(u => {
            userMap[String(u[1]).trim()] = { name: u[0], id: u[1], com: u[2], group: u[3], part: u[4], role: u[5] };
        });

        const allCourses = courseSheet.getDataRange().getDisplayValues();
        const courseMap = {};
        allCourses.slice(1).forEach(c => courseMap[String(c[0]).trim()] = c[1]);

        const currentAssigns = assignSheet.getDataRange().getDisplayValues().slice(1);
        const assignedSet = new Set(currentAssigns.map(r => `${String(r[0]).trim()}_${String(r[2]).trim()}`));

        const newRows = [];
        let count = 0;

        userIds.forEach(uid => {
            const uStr = String(uid).trim();
            const user = userMap[uStr];
            if (user) {
                const isAdmin = (String(user.role).toLowerCase() === 'admin') || (uStr.toLowerCase() === 'admin');
                if (!isAdmin && String(user.role).toLowerCase() === 'user') {
                    courseIds.forEach(cid => {
                        const cStr = String(cid).trim();
                        const key = `${cStr}_${uStr}`;
                        if (courseMap[cStr] && !assignedSet.has(key)) {
                            newRows.push([
                                cStr,
                                courseMap[cStr],
                                user.id,
                                user.name,
                                user.com,
                                user.group,
                                user.part,
                                '', '', ''
                            ]);
                            assignedSet.add(key);
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
    } catch (e) {
        return { error: e.message };
    }
}

function processNewRegisters(userIds, courseIds, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const assignRes = assignCoursesToUsers(userIds, courseIds, adminToken);
        if (assignRes.error) return assignRes;

        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const idSet = new Set(userIds.map(id => String(id).trim()));

        for (let i = 1; i < data.length; i++) {
            if (idSet.has(String(data[i][1]).trim())) {
                sheet.getRange(i + 1, 7).setValue('Active');
            }
        }
        return { success: true, count: assignRes.count };
    } catch (e) {
        return { error: e.message };
    }
}

function importUsers(userList, courseIds, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const sheet = getDB().getSheetByName('Users');
        const data = sheet.getDataRange().getDisplayValues();
        const existingIds = new Set(data.slice(1).map(r => String(r[1]).trim()));

        let addedCount = 0;
        let errors = [];
        let newRows = [];
        let newUserIds = [];

        userList.forEach(u => {
            if (!u.id || !u.name) return;
            const idStr = String(u.id).trim();
            if (existingIds.has(idStr)) {
                errors.push("Trùng mã nhân viên: " + idStr);
            } else {
                newRows.push([u.name, idStr, u.com || '', u.group || '', u.part || '', 'user', 'Active']);
                existingIds.add(idStr);
                newUserIds.push(idStr);
                addedCount++;
            }
        });

        if (newRows.length > 0) {
            const startRow = sheet.getLastRow() + 1;
            sheet.getRange(startRow, 2, newRows.length, 1).setNumberFormat("@");
            sheet.getRange(startRow, 1, newRows.length, 7).setValues(newRows);

            if (courseIds && courseIds.length > 0) {
                assignCoursesToUsers(newUserIds, courseIds, adminToken);
            }
        }

        return { success: true, added: addedCount, errors: errors };
    } catch (e) {
        return { error: "Lỗi import: " + e.message };
    }
}

function getUserStats(adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const users = getDB().getSheetByName('Users').getDataRange().getDisplayValues().slice(1);
        const stats = {};
        let total = 0;

        users.forEach(u => {
            const isAdm = (String(u[5]).toLowerCase() === 'admin') || (String(u[1]).toLowerCase() === 'admin');
            if (!isAdm) {
                const part = u[4] || 'Chưa phân loại';
                stats[part] = (stats[part] || 0) + 1;
                total++;
            }
        });

        return { stats: stats, total: total };
    } catch (e) {
        return { error: "Lỗi thống kê: " + e.message };
    }
}

function deleteUsersByScope(scope, value, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const ss = getDB();
        const userSheet = ss.getSheetByName('Users');
        const uData = userSheet.getDataRange().getDisplayValues();
        if (uData.length <= 1) return { success: true, count: 0 };

        const header = uData[0];
        const remainingUsers = [header];
        const deletedIDs = new Set();
        let deletedCount = 0;

        for (let i = 1; i < uData.length; i++) {
            const u = uData[i];
            const uID = String(u[1]).trim();
            const uPart = String(u[4]).trim();
            const uRole = String(u[5]).trim().toLowerCase();

            if (uRole === 'admin' || uID.toLowerCase() === 'admin') {
                remainingUsers.push(u);
                continue;
            }

            let shouldDelete = false;
            if (scope === 'all') {
                shouldDelete = true;
            } else if (scope === 'part' && uPart === String(value).trim()) {
                shouldDelete = true;
            }

            if (shouldDelete) {
                deletedIDs.add(uID);
                deletedCount++;
            } else {
                remainingUsers.push(u);
            }
        }

        if (deletedCount > 0) {
            userSheet.getRange(1, 1, uData.length, uData[0].length).clearContent();
            userSheet.getRange(1, 1, remainingUsers.length, remainingUsers[0].length).setValues(remainingUsers);

            const assignSheet = ss.getSheetByName('Assignments');
            if (assignSheet) {
                const aData = assignSheet.getDataRange().getDisplayValues();
                if (aData.length > 1) {
                    const remainingAssigns = [aData[0]];
                    for (let j = 1; j < aData.length; j++) {
                        const aUid = String(aData[j][2]).trim();
                        const isIncomplete = aData[j][7] === '' && aData[j][9] === '';
                        if (!(deletedIDs.has(aUid) && isIncomplete)) {
                            remainingAssigns.push(aData[j]);
                        }
                    }
                    assignSheet.getRange(1, 1, aData.length, aData[0].length).clearContent();
                    assignSheet.getRange(1, 1, remainingAssigns.length, remainingAssigns[0].length).setValues(remainingAssigns);
                }
            }
        }

        return { success: true, count: deletedCount };
    } catch (e) {
        return { error: "Lỗi xóa dữ liệu: " + e.message };
    }
}

// --- USER ASSIGNMENTS & ULTRA LOW-LATENCY SIGNATURE WORKFLOW ---

/**
 * Tải danh sách khóa học kèm nội dung chi tiết (Zero-Latency Modal):
 * Giúp client mở modal xem nội dung bài học ngay tức thì mà không cần gọi API getCourseDetails
 */
function getMyCourses(userID) {
    try {
        const db = getDB();
        const assignData = db.getSheetByName('Assignments').getDataRange().getDisplayValues();
        const courseData = db.getSheetByName('Courses').getDataRange().getDisplayValues();
        const uTarget = String(userID).trim();

        const courseMap = {};
        for (let i = 1; i < courseData.length; i++) {
            courseMap[String(courseData[i][0]).trim()] = courseData[i];
        }

        const assigns = assignData.slice(1);
        const myAssigns = assigns.filter(r => String(r[2]).trim() === uTarget && r[7] === '' && r[9] === '');

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return myAssigns.map(a => {
            const cid = String(a[0]).trim();
            const courseInfo = courseMap[cid];
            if (!courseInfo) return null;

            const startDate = parseDateFlexible(courseInfo[2]);
            const endDate = parseDateFlexible(courseInfo[3]);

            const isCourseOpen = (String(courseInfo[6]).toLowerCase() === 'true');
            let status = 'Opening';

            if (endDate && today > endDate) status = 'Ended';
            else if (startDate && today < startDate) status = 'Plan';

            const canSign = (status === 'Opening' && isCourseOpen);

            return {
                id: cid,
                name: a[1],
                start: courseInfo[2],
                end: courseInfo[3],
                status: status,
                canSign: canSign,
                content: courseInfo[4] || '' // Trả về sẵn nội dung để client mở modal 0ms!
            };
        }).filter(Boolean);
    } catch (e) {
        return [];
    }
}

/**
 * Xác nhận đào tạo - Có Script Lock ngăn chặn Race Condition khi hàng trăm nhân viên ký cùng lúc
 */
function confirmTraining(userID, courseID, sign) {
    if (!sign || sign.length < 500) {
        return { success: false, message: "Dữ liệu chữ ký không hợp lệ hoặc bị trống!" };
    }

    const lock = LockService.getScriptLock();
    try {
        // Đợi tối đa 10 giây để lấy quyền ghi độc quyền
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang xử lý nhiều lượt ký đồng thời. Vui lòng thử lại sau 3 giây!" };
    }

    try {
        const uTarget = String(userID).trim();
        const cTarget = String(courseID).trim();
        const db = getDB();
        const courseSheet = db.getSheetByName('Courses');
        const courseData = courseSheet.getDataRange().getDisplayValues();

        const course = courseData.find(c => String(c[0]).trim() === cTarget);
        if (!course) return { success: false, message: "Không tìm thấy khóa đào tạo!" };

        if (String(course[6]).toLowerCase() !== 'true') {
            return { success: false, message: "Khóa đào tạo hiện đang khóa chức năng ký!" };
        }

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const start = parseDateFlexible(course[2]);
        const end = parseDateFlexible(course[3]);

        if ((start && today < start) || (end && today > end)) {
            return { success: false, message: "Thời hạn khóa học không nằm trong khoảng được ký!" };
        }

        const assignSheet = db.getSheetByName('Assignments');
        const aData = assignSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < aData.length; i++) {
            if (String(aData[i][2]).trim() === uTarget && String(aData[i][0]).trim() === cTarget) {
                // Kiểm tra nếu đã ký rồi để tránh tạo file rác trong Drive
                if (aData[i][7] !== '') {
                    return { success: true, message: "Khóa học này đã được ký xác nhận trước đó!" };
                }

                const signRef = saveSignatureToDrive(cTarget, uTarget, sign);
                assignSheet.getRange(i + 1, 8, 1, 2).setValues([[signRef, new Date()]]);
                return { success: true };
            }
        }

        return { success: false, message: "Không tìm thấy lượt phân công tương ứng để cập nhật!" };
    } catch (err) {
        return { success: false, message: "Lỗi xử lý xác nhận: " + err.message };
    } finally {
        lock.releaseLock();
    }
}

// --- QUẢN LÝ KHÓA HỌC (COURSES) ---

function updateCourseSignStatus(courseID, isOpen, adminToken) {
    if (!verifyAdminToken(adminToken)) return false;
    try {
        clearDashboardCache();
        const sheet = getDB().getSheetByName('Courses');
        const data = sheet.getDataRange().getDisplayValues();
        const cTarget = String(courseID).trim();

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === cTarget) {
                sheet.getRange(i + 1, 7).setNumberFormat("@").setValue(isOpen ? 'true' : 'false');
                return true;
            }
        }
        return false;
    } catch (e) {
        return false;
    }
}

function saveCourseWithTarget(courseData, targetType, targetValue, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const ss = getDB();
        const courseSheet = ss.getSheetByName('Courses');
        const assignSheet = ss.getSheetByName('Assignments');
        const userSheet = ss.getSheetByName('Users');

        const isNew = !courseData.id;
        if (isNew) courseData.id = 'TRN-' + new Date().getTime();

        const courseDataRow = courseSheet.getDataRange().getDisplayValues();
        let existingCourseIndex = -1;

        for (let i = 1; i < courseDataRow.length; i++) {
            if (String(courseDataRow[i][0]).trim() === String(courseData.id).trim()) {
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
            let idList = Array.isArray(targetValue) ? targetValue.map(id => String(id).trim()) : targetValue.split(',').map(id => id.trim());
            targetUsers = users.slice(1).filter(u => idList.includes(String(u[1]).trim()));
        }

        targetUsers = targetUsers.filter(u => {
            const role = String(u[5] || '').toLowerCase();
            const uid = String(u[1] || '').toLowerCase();
            return role === 'user' && uid !== 'admin';
        });

        if (targetUsers.length > 0) {
            const rows = targetUsers.map(u => [
                courseData.id,
                courseData.name,
                u[1],
                u[0],
                u[2],
                u[3],
                u[4],
                '', '', ''
            ]);
            assignSheet.getRange(assignSheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
        }
        return { success: true, count: targetUsers.length, action: 'created' };
    } catch (e) {
        return { error: e.message };
    }
}

function deleteCourse(courseId, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        clearDashboardCache();
        const ss = getDB();
        const courseSheet = ss.getSheetByName('Courses');
        const assignSheet = ss.getSheetByName('Assignments');
        const cTarget = String(courseId).trim();

        const cData = courseSheet.getDataRange().getDisplayValues();
        for (let i = 1; i < cData.length; i++) {
            if (String(cData[i][0]).trim() === cTarget) {
                courseSheet.deleteRow(i + 1);
                break;
            }
        }

        const aData = assignSheet.getDataRange().getDisplayValues();
        if (aData.length > 1) {
            const safeRows = aData.filter(r => String(r[0]).trim() !== cTarget);
            if (safeRows.length < aData.length) {
                assignSheet.getRange(1, 1, aData.length, aData[0].length).clearContent();
                assignSheet.getRange(1, 1, safeRows.length, safeRows[0].length).setValues(safeRows);
            }
        }

        return true;
    } catch (e) {
        return { error: e.message };
    }
}

function getCourseStatistics() {
    try {
        const ss = getDB();
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const assignMap = {};
        assignments.forEach(a => {
            const cid = String(a[0]).trim();
            if (!assignMap[cid]) assignMap[cid] = { total: 0, completed: 0 };
            assignMap[cid].total++;
            if (a[7] !== '' || a[9] !== '') assignMap[cid].completed++;
        });

        return courses.map(c => {
            const cid = String(c[0]).trim();
            const stats = assignMap[cid] || { total: 0, completed: 0 };
            let percent = stats.total > 0 ? Math.round((stats.completed / stats.total) * 100) : 0;

            const startDate = parseDateFlexible(c[2]);
            const endDate = parseDateFlexible(c[3]);
            let status = '';

            const isSignEnabled = (String(c[6]).toLowerCase() === 'true');

            if (percent === 100 && stats.total > 0) status = 'Complete';
            else if (endDate && today > endDate) status = 'Pending';
            else if (startDate && today < startDate) status = 'Plan';
            else status = 'Opening';

            return {
                id: cid, name: c[1], start: c[2], end: c[3], content: c[4],
                status: status, percent: percent, pendingCount: stats.total - stats.completed,
                signEnabled: isSignEnabled
            };
        }).reverse();
    } catch (e) {
        return { error: "Lỗi tải khóa học: " + e.message };
    }
}

function getCourseParticipants(courseID, adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return [];
    try {
        const cTarget = String(courseID).trim();
        const data = getDB().getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        return data.filter(r => String(r[0]).trim() === cTarget).map(r => ({
            uid: r[2], name: r[3], com: r[4], group: r[5], part: r[6],
            signed: r[7] !== '',
            signTime: r[8],
            exception: r[9]
        }));
    } catch (e) {
        return [];
    }
}

function updateException(courseID, userID, reason, adminToken) {
    if (!verifyAdminToken(adminToken)) return false;
    try {
        const sheet = getDB().getSheetByName('Assignments');
        const data = sheet.getDataRange().getDisplayValues();
        const cTarget = String(courseID).trim();
        const uTarget = String(userID).trim();

        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]).trim() === cTarget && String(data[i][2]).trim() === uTarget) {
                sheet.getRange(i + 1, 10).setValue(reason);
                return true;
            }
        }
        return false;
    } catch (e) {
        return false;
    }
}

// --- EXPORT BÁO CÁO NHẮC NHỞ & EXCEL ---

function exportConsolidatedReminder(adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const ss = getDB();
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const targetCourses = {};
        const courseMap = {};

        courses.forEach(c => {
            const cid = String(c[0]).trim();
            const start = parseDateFlexible(c[2]);
            const end = parseDateFlexible(c[3]);
            courseMap[cid] = { name: c[1], end: end };

            if (start && today >= start) {
                targetCourses[cid] = c[1];
            }
        });

        const reportData = [];
        assignments.forEach(a => {
            const cID = String(a[0]).trim();
            const cName = targetCourses[cID];
            if (cName && a[7] === '' && a[9] === '') {
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

        if (reportData.length === 0) return { error: "Hiện không có dữ liệu nhắc nhở nào cần xuất!" };

        const folder = getDataFolder();
        const tempSS = SpreadsheetApp.create("Temp_Export_" + Date.now());
        const tempFile = DriveApp.getFileById(tempSS.getId());
        tempFile.moveTo(folder);

        const sheet = tempSS.getActiveSheet();
        const headers = ["Khóa đào tạo", "Họ và tên", "Mã nhân viên", "Company", "Group", "Part", "Trạng thái"];
        sheet.appendRow(headers);

        const rows = reportData.map(r => {
            const cEnd = courseMap[r.cID] ? courseMap[r.cID].end : null;
            const status = (cEnd && today > cEnd) ? "Quá hạn" : "Chưa hoàn thành";
            return [targetCourses[r.cID], r.name, r.uID, r.com, r.group, r.part, status];
        });

        sheet.getRange(2, 1, rows.length, 7).setValues(rows);

        const fullRange = sheet.getDataRange();
        fullRange.setFontFamily("Arial");
        fullRange.setFontSize(11);
        sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#F1F5F9").setHorizontalAlignment("center");
        fullRange.setBorder(true, true, true, true, true, true);
        sheet.autoResizeColumns(1, 7);

        SpreadsheetApp.flush();

        const url = "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
        const token = ScriptApp.getOAuthToken();
        const response = UrlFetchApp.fetch(url, {
            headers: { 'Authorization': 'Bearer ' + token },
            muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
            return { error: "Lỗi kết xuất Excel: " + response.getResponseCode() };
        }

        const blob = response.getBlob().setName("IQC_NhacNho_DaoTao.xlsx");
        const base64 = Utilities.base64Encode(blob.getBytes());

        try { tempFile.setTrashed(true); } catch (e) { }

        return { filename: "IQC_NhacNho_DaoTao.xlsx", base64: base64 };
    } catch (e) {
        return { error: "Lỗi xuất file: " + e.message };
    }
}

function _forceScope() { DriveApp.getRootFolder(); }

// --- AUTOMATION TRIGGERS & EMAIL REPORTING ---

function setupTriggers() {
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => ScriptApp.deleteTrigger(t));

    ScriptApp.newTrigger('sendMorningReport')
        .timeBased()
        .everyDays(1)
        .atHour(7)
        .create();

    ScriptApp.newTrigger('sendAfternoonReport')
        .timeBased()
        .everyDays(1)
        .atHour(15)
        .create();

    ScriptApp.newTrigger('checkCourseCompletion')
        .timeBased()
        .everyDays(1)
        .atHour(6)
        .create();
}

function sendMorningReport() {
    const day = new Date().getDay();
    if (day >= 1 && day <= 5) {
        sendIncompleteReport("Sáng (07:00)");
    }
}

function sendAfternoonReport() {
    const day = new Date().getDay();
    if (day >= 1 && day <= 5) {
        sendIncompleteReport("Chiều (15:00)");
    }
}

function sendIncompleteReport(timeLabel) {
    try {
        const ss = getDB();
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const courseMap = {};
        courses.forEach(c => {
            const end = parseDateFlexible(c[3]);
            courseMap[String(c[0]).trim()] = { name: c[1], end: end };
        });

        const incompleteList = [];
        assignments.forEach(a => {
            if (a[7] === '' && a[9] === '') {
                const cInfo = courseMap[String(a[0]).trim()];
                if (cInfo && cInfo.end && today <= cInfo.end) {
                    incompleteList.push({
                        cName: cInfo.name,
                        uName: a[3],
                        uID: a[2],
                        uPart: a[6]
                    });
                }
            }
        });

        if (incompleteList.length === 0) return;

        let html = `
        <div style="font-family: Arial, sans-serif; color: #1E293B; max-width: 700px; margin: auto; border: 1px solid #E2E8F0; border-radius: 8px; padding: 24px;">
            <div style="border-bottom: 2px solid #2563EB; padding-bottom: 12px; margin-bottom: 16px;">
                <h2 style="color: #1E293B; margin: 0; font-size: 18px;">IQC NEXUS | BÁO CÁO ĐÀO TẠO CHƯA HOÀN THÀNH - ${timeLabel}</h2>
            </div>
            <p style="font-size: 14px;">Kính gửi Anh/Chị Quản trị viên,</p>
            <p style="font-size: 14px;">Hệ thống ghi nhận tổng số <b>${incompleteList.length}</b> trường hợp nhân viên chưa thực hiện ký xác nhận các khóa học đang diễn ra:</p>
            <table border="1" cellpadding="8" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 13px; border-color: #CBD5E1;">
                <tr style="background-color: #F8FAFC; text-align: left;">
                    <th>Khóa học</th><th>Họ và tên</th><th>Mã NV</th><th>Bộ phận</th>
                </tr>`;

        incompleteList.forEach(item => {
            html += `<tr>
                <td>${item.cName}</td>
                <td><b>${item.uName}</b></td>
                <td>${item.uID}</td>
                <td>${item.uPart}</td>
            </tr>`;
        });

        html += `
            </table>
            <p style="margin-top: 16px; font-size: 12px; color: #64748B;"><i>Email này được tạo tự động bởi hệ thống IQC Nexus. Vui lòng xem bảng đính kèm để biết thông tin chi tiết.</i></p>
        </div>`;

        let attachmentBlob = null;
        let tempFile = null;

        try {
            const folder = getDataFolder();
            const tempSS = SpreadsheetApp.create("Temp_Report_" + Date.now());
            tempFile = DriveApp.getFileById(tempSS.getId());
            tempFile.moveTo(folder);

            const sheet = tempSS.getActiveSheet();
            sheet.appendRow(["Tên Khóa Học", "Họ và Tên", "Mã Nhân Viên", "Bộ Phận"]);
            const rows = incompleteList.map(item => [item.cName, item.uName, item.uID, item.uPart]);

            if (rows.length > 0) {
                sheet.getRange(2, 1, rows.length, 4).setValues(rows);
            }

            sheet.getDataRange().setFontFamily("Arial").setBorder(true, true, true, true, true, true);
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
            console.error("Lỗi tạo file đính kèm: " + err.message);
        }

        const recipients = getProp('RECIPIENT_EMAILS', RECIPIENT_EMAILS_DEFAULT);
        const emailOptions = {
            to: recipients,
            subject: `[IQC Nexus] Báo cáo đào tạo chưa hoàn thành - ${timeLabel}`,
            htmlBody: html
        };

        if (attachmentBlob) {
            emailOptions.attachments = [attachmentBlob];
        }

        MailApp.sendEmail(emailOptions);

        if (tempFile) {
            try { tempFile.setTrashed(true); } catch (e) { }
        }
    } catch (e) {
        console.error("Lỗi gửi báo cáo định kỳ: " + e.message);
    }
}

function checkCourseCompletion() {
    try {
        const ss = getDB();
        const courses = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        const yesterday = new Date();
        yesterday.setDate(yesterday.getDate() - 1);
        yesterday.setHours(0, 0, 0, 0);

        courses.forEach(c => {
            const endDate = parseDateFlexible(c[3]);
            if (endDate && endDate.getTime() === yesterday.getTime()) {
                sendCourseFinalReport(c, assignments);
            }
        });
    } catch (e) {
        console.error("Lỗi kiểm tra hoàn thành khóa học: " + e.message);
    }
}

function sendCourseFinalReport(courseRow, allAssignments) {
    const courseId = String(courseRow[0]).trim();
    const courseName = courseRow[1];

    const courseAssigns = allAssignments.filter(a => String(a[0]).trim() === courseId);
    if (courseAssigns.length === 0) return;

    const total = courseAssigns.length;
    const completed = courseAssigns.filter(a => a[7] !== '' || a[9] !== '').length;
    const incomplete = courseAssigns.filter(a => a[7] === '' && a[9] === '');
    const percent = Math.round((completed / total) * 100);

    let html = `
    <div style="font-family: Arial, sans-serif; color: #1E293B; max-width: 700px; margin: auto; border: 1px solid #E2E8F0; border-radius: 8px; padding: 24px;">
        <h2 style="color: #1E293B; margin-top: 0; font-size: 18px; border-bottom: 2px solid #10B981; padding-bottom: 10px;">
            IQC NEXUS | BÁO CÁO TỔNG KẾT KHÓA HỌC: ${courseName}
        </h2>
        <p><b>Thời gian kết thúc:</b> ${courseRow[3]}</p>
        <p><b>Tổng số học viên:</b> ${total}</p>
        <p><b>Đã hoàn thành:</b> <span style="color: #10B981; font-weight: bold;">${completed} (${percent}%)</span></p>
        <p><b>Chưa hoàn thành:</b> <span style="color: #EF4444; font-weight: bold;">${incomplete.length}</span></p>
    `;

    if (incomplete.length > 0) {
        html += `
        <h4 style="margin-top: 20px;">Danh sách nhân viên chưa hoàn thành:</h4>
        <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 13px; border-color: #CBD5E1;">
            <tr style="background: #F8FAFC;"><th>Nhân viên</th><th>Mã NV</th><th>Bộ phận</th></tr>`;
        incomplete.forEach(row => {
            html += `<tr><td>${row[3]}</td><td>${row[2]}</td><td>${row[6]}</td></tr>`;
        });
        html += `</table>`;
    }

    html += `<p style="margin-top: 16px; font-size: 12px; color: #64748B;"><i>Email này được gửi tự động từ hệ thống IQC Nexus.</i></p></div>`;

    const recipients = getProp('RECIPIENT_EMAILS', RECIPIENT_EMAILS_DEFAULT);
    MailApp.sendEmail({
        to: recipients,
        subject: `[IQC Nexus] Tổng kết khóa đào tạo: ${courseName}`,
        htmlBody: html
    });
}

function exportPendingList(courseId) {
    const cTarget = String(courseId).trim();
    const ss = getDB();
    const assigns = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
    let pending = assigns.filter(a => String(a[0]).trim() === cTarget && a[7] === '' && a[9] === '');
    return pending.map(a => [a[3], a[2], a[4], a[5], a[6]]);
}

// --- XUẤT BÁO CÁO PDF ---

function exportCourseReport(courseID, adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const cTarget = String(courseID).trim();
        const ss = getDB();
        const courseData = ss.getSheetByName('Courses').getDataRange().getDisplayValues().slice(1);
        const assignData = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        const course = courseData.find(c => String(c[0]).trim() === cTarget);
        if (!course) return { error: "Không tìm thấy khóa học!" };

        const cName = String(course[1]).toUpperCase();
        const cTime = `${course[2]} - ${course[3]}`;
        const cContent = course[4];

        const participants = assignData.filter(a =>
            String(a[0]).trim() === cTarget && (a[7] !== '' || a[9] !== '')
        );

        if (participants.length === 0) {
            return { error: "Hiện chưa có học viên nào hoàn thành để xuất báo cáo PDF!" };
        }

        let html = `
        <html>
        <head>
            <style>
                body { font-family: 'Times New Roman', serif; padding: 24px; font-size: 11pt; color: #000; }
                h1 { font-size: 16pt; text-align: center; margin-bottom: 6px; text-transform: uppercase; font-weight: bold; }
                .meta { text-align: center; font-size: 11pt; margin-bottom: 16px; font-style: italic; }
                .content-box { margin-bottom: 16px; padding: 10px; border: 1px solid #ccc; font-size: 10.5pt; }
                table { width: 100%; border-collapse: collapse; font-size: 10pt; }
                th, td { border: 1px solid #000; padding: 6px; vertical-align: middle; text-align: center; }
                th { font-weight: bold; background-color: #f2f2f2; }
                .left { text-align: left; }
                img.sign { max-height: 38px; max-width: 90px; object-fit: contain; }
            </style>
        </head>
        <body>
            <h1>DANH SÁCH XÁC NHẬN ĐÀO TẠO</h1>
            <h2 style="font-size: 13pt; text-align: center; margin-top: 0;">KHÓA: ${cName}</h2>
            <div class="meta">Thời gian đào tạo: ${cTime}</div>
            <div class="content-box"><strong>Nội dung đào tạo:</strong> ${cContent}</div>

            <table>
                <thead>
                    <tr>
                        <th style="width: 5%">STT</th>
                        <th style="width: 22%">Họ và tên</th>
                        <th style="width: 11%">Mã NV</th>
                        <th style="width: 12%">Company</th>
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
            const signImg = resolveSignatureImageBase64(p[7]);
            const time = p[8] ? new Date(p[8]).toLocaleString('vi-VN') : '';
            const exc = p[9];

            let signCell = '';
            let metaCell = '';

            if (exc !== '') {
                signCell = '<span style="color:#d97706;font-style:italic">Ngoại lệ</span>';
                metaCell = exc;
            } else {
                if (signImg.startsWith('data:image')) {
                    signCell = `<img src="${signImg}" class="sign" />`;
                } else {
                    signCell = '<span>Đã ký</span>';
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

        const dataFolder = getDataFolder();
        const blob = Utilities.newBlob(html, MimeType.HTML, "temp.html");
        const tempFile = dataFolder.createFile(blob);
        const pdfBlob = tempFile.getAs(MimeType.PDF);
        const finalBase64 = Utilities.base64Encode(pdfBlob.getBytes());

        tempFile.setTrashed(true);

        return { filename: `BaoCao_IQC_${cName}.pdf`, base64: finalBase64 };
    } catch (e) {
        return { error: "Lỗi tạo báo cáo PDF: " + e.message };
    }
}
