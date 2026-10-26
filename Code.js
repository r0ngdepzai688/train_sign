/**
 * IQC NEXUS - HỆ THỐNG QUẢN LÝ VÀ KÝ XÁC NHẬN ĐÀO TẠO
 * Phiên bản: 2.2 (High-Performance Caching & Optimized Spreadsheet I/O)
 * Tác giả: Nguyễn Hải Đường (hai.duong)
 * Slogan: Silence the ads. Let quality make the noise.
 */

// --- CẤU HÌNH HỆ THỐNG & V8 IN-MEMORY MULTI-TIER CACHE ---
const ADMIN_PASS_DEFAULT = "@nhd16488";
const USER_DEFAULT_PASS_DEFAULT = "@iqc6368";
const SPREADSHEET_ID_DEFAULT = "15K3FetZvSMPRKLNp9I7kyPLCeAIpreWNeeG7uerW3wk";
const RECIPIENT_EMAILS_DEFAULT = 'hien.pt1@samsung.com,qcteam.iqctraining@gmail.com,dinh.hnhung@samsung.com,dang.ngoc@samsung.com,nhung.tt2@samsung.com,an.pthi@samsung.com,haiduong17.nguyen@gmail.com';

const _MEM_CACHE = {
    hydrated: false,
    db: null,
    props: null,
    usersData: null,
    coursesData: null,
    assignIndex: null,
    courseStats: null,
    adminFilters: null,
    availableCourses: null,
    authTokens: {}
};

function hydrateMemoryFromScriptCache() {
    if (_MEM_CACHE.hydrated) return;
    _MEM_CACHE.hydrated = true;
    try {
        const cache = CacheService.getScriptCache();
        const keys = [
            'sys_props_v2',
            'users_raw_v2',
            'courses_raw_data',
            'assign_idx_v2',
            'course_stats_cache',
            'admin_filters',
            'available_courses'
        ];
        const batch = cache.getAll(keys) || {};
        if (batch['sys_props_v2'] && !_MEM_CACHE.props) _MEM_CACHE.props = JSON.parse(batch['sys_props_v2']);
        if (batch['users_raw_v2'] && !_MEM_CACHE.usersData) _MEM_CACHE.usersData = JSON.parse(batch['users_raw_v2']);
        if (batch['courses_raw_data'] && !_MEM_CACHE.coursesData) _MEM_CACHE.coursesData = JSON.parse(batch['courses_raw_data']);
        if (batch['assign_idx_v2'] && !_MEM_CACHE.assignIndex) _MEM_CACHE.assignIndex = JSON.parse(batch['assign_idx_v2']);
        if (batch['course_stats_cache'] && !_MEM_CACHE.courseStats) _MEM_CACHE.courseStats = JSON.parse(batch['course_stats_cache']);
        if (batch['admin_filters'] && !_MEM_CACHE.adminFilters) _MEM_CACHE.adminFilters = JSON.parse(batch['admin_filters']);
        if (batch['available_courses'] && !_MEM_CACHE.availableCourses) _MEM_CACHE.availableCourses = JSON.parse(batch['available_courses']);
    } catch (e) { }
}

function getProp(key, fallback) {
    hydrateMemoryFromScriptCache();
    if (!_MEM_CACHE.props) {
        try {
            const rawProps = PropertiesService.getScriptProperties().getProperties() || {};
            _MEM_CACHE.props = rawProps;
            try {
                CacheService.getScriptCache().put('sys_props_v2', JSON.stringify(rawProps), 3600);
            } catch (e) { }
        } catch (e) {
            _MEM_CACHE.props = {};
        }
    }
    const val = _MEM_CACHE.props[key];
    return (val !== null && val !== undefined && val !== '') ? val : fallback;
}

function getDB() {
    if (_MEM_CACHE.db) return _MEM_CACHE.db;
    const sheetId = getProp('SPREADSHEET_ID', SPREADSHEET_ID_DEFAULT);
    _MEM_CACHE.db = SpreadsheetApp.openById(sheetId);
    return _MEM_CACHE.db;
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
    if (_MEM_CACHE.authTokens[token] === 'admin') return true;
    try {
        const cached = CacheService.getScriptCache().get('AUTH_' + token);
        if (cached === 'admin') {
            _MEM_CACHE.authTokens[token] = 'admin';
            return true;
        }
        return false;
    } catch (e) {
        return false;
    }
}

/**
 * Khởi động nóng toàn bộ Engine & nạp sẵn Index vào V8 RAM + CacheService
 * Được gọi tự động ngầm ngay khi người dùng vừa mở màn hình Đăng nhập
 */
function prewarmSystem() {
    try {
        hydrateMemoryFromScriptCache();
        getProp('ADMIN_PASS', ADMIN_PASS_DEFAULT);
        getCachedUsersData();
        getCachedCourses();
        getCachedAssignmentsIndex();
        getCourseStatistics();
        getAdminFilters();
        getAvailableCourses();
        return { warmed: true, ts: Date.now() };
    } catch (e) {
        return { warmed: false };
    }
}

/**
 * Đăng nhập siêu tốc (Zero-Roundtrip + O(1) RAM Index Lookup)
 */
function checkLogin(user, pass) {
    try {
        hydrateMemoryFromScriptCache();
        const adminPass = getProp('ADMIN_PASS', ADMIN_PASS_DEFAULT);
        const userDefaultPass = getProp('USER_DEFAULT_PASS', USER_DEFAULT_PASS_DEFAULT);
        const uTarget = String(user || '').trim();

        // 1. Quản trị viên
        if (uTarget === 'admin' && pass === adminPass) {
            const token = Utilities.getUuid() + '-' + Date.now();
            _MEM_CACHE.authTokens[token] = 'admin';
            try {
                CacheService.getScriptCache().put('AUTH_' + token, 'admin', 21600); // 6 giờ
            } catch (e) { }

            const dashboardData = getAdminDashboardData(token, true);
            return {
                role: 'admin',
                user: 'Quản trị viên',
                userID: 'admin',
                token: token,
                initialDashboard: dashboardData
            };
        }

        // 2. Nhân viên IQC (Đọc từ Cached Users Index - 0ms Spreadsheet I/O khi Warm)
        const userData = getCachedUsersData();
        if (!userData || userData.length === 0) return { error: "Không tìm thấy bảng Users!" };

        for (let i = 1; i < userData.length; i++) {
            const uID = String(userData[i][1] || '').trim();
            if (uID === uTarget && pass === userDefaultPass) {
                if (userData[i][6] === 'Inactive') return { error: "Tài khoản của bạn đã bị khóa!" };

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

// --- UTILITIES: DATE PARSER & ADVANCED CACHE MANAGEMENT ---

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

/**
 * Cache bảng Users trong V8 RAM + CacheService (TTL 30 phút)
 */
function getCachedUsersData() {
    hydrateMemoryFromScriptCache();
    if (_MEM_CACHE.usersData) return _MEM_CACHE.usersData;
    try {
        const cached = CacheService.getScriptCache().get('users_raw_v2');
        if (cached) {
            _MEM_CACHE.usersData = JSON.parse(cached);
            return _MEM_CACHE.usersData;
        }
    } catch (e) { }

    const ss = getDB();
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    _MEM_CACHE.usersData = data;
    try {
        const serialized = JSON.stringify(data);
        if (serialized.length < 95000) {
            CacheService.getScriptCache().put('users_raw_v2', serialized, 1800);
        }
    } catch (e) { }
    return data;
}

/**
 * Cache trung gian danh sách khóa học:
 * Tránh đọc bảng Courses liên tục trong mỗi request
 */
function getCachedCourses() {
    hydrateMemoryFromScriptCache();
    if (_MEM_CACHE.coursesData) return _MEM_CACHE.coursesData;
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('courses_raw_data');
        if (cached) {
            _MEM_CACHE.coursesData = JSON.parse(cached);
            return _MEM_CACHE.coursesData;
        }
    } catch (e) { }

    const ss = getDB();
    const sheet = ss.getSheetByName('Courses');
    if (!sheet) return [];
    const data = sheet.getDataRange().getDisplayValues();
    _MEM_CACHE.coursesData = data;
    try {
        const serialized = JSON.stringify(data);
        if (serialized.length < 95000) {
            CacheService.getScriptCache().put('courses_raw_data', serialized, 1800);
        }
    } catch (e) { }
    return data;
}

/**
 * Cache chỉ mục phân bổ gọn nhẹ (statsByCourse + pendingByUser):
 * Giúp checkLogin(user) & getCourseStatistics() đạt 0ms Spreadsheet I/O
 */
function getCachedAssignmentsIndex() {
    hydrateMemoryFromScriptCache();
    if (_MEM_CACHE.assignIndex) return _MEM_CACHE.assignIndex;
    try {
        const cached = CacheService.getScriptCache().get('assign_idx_v2');
        if (cached) {
            _MEM_CACHE.assignIndex = JSON.parse(cached);
            return _MEM_CACHE.assignIndex;
        }
    } catch (e) { }

    const ss = getDB();
    const sheet = ss.getSheetByName('Assignments');
    const rows = sheet ? sheet.getDataRange().getDisplayValues().slice(1) : [];

    const statsByCourse = {};
    const pendingByUser = {};

    for (let i = 0; i < rows.length; i++) {
        const r = rows[i];
        const cid = String(r[0] || '').trim();
        const uid = String(r[2] || '').trim();
        if (!cid) continue;

        if (!statsByCourse[cid]) statsByCourse[cid] = { total: 0, completed: 0 };
        statsByCourse[cid].total++;

        const isDone = (r[7] !== '' || r[9] !== '');
        if (isDone) {
            statsByCourse[cid].completed++;
        } else if (uid) {
            if (!pendingByUser[uid]) pendingByUser[uid] = [];
            pendingByUser[uid].push(cid);
        }
    }

    const indexObj = { statsByCourse: statsByCourse, pendingByUser: pendingByUser };
    _MEM_CACHE.assignIndex = indexObj;
    try {
        const serialized = JSON.stringify(indexObj);
        if (serialized.length < 95000) {
            CacheService.getScriptCache().put('assign_idx_v2', serialized, 1800);
        }
    } catch (e) { }
    return indexObj;
}

function clearDashboardCache() {
    _MEM_CACHE.usersData = null;
    _MEM_CACHE.coursesData = null;
    _MEM_CACHE.assignIndex = null;
    _MEM_CACHE.courseStats = null;
    _MEM_CACHE.adminFilters = null;
    _MEM_CACHE.availableCourses = null;
    try {
        const cache = CacheService.getScriptCache();
        const keysToRemove = [
            'users_raw_v2',
            'assign_idx_v2',
            'admin_filters',
            'available_courses',
            'courses_raw_data',
            'course_stats_cache'
        ];
        if (typeof cache.removeAll === 'function') {
            cache.removeAll(keysToRemove);
        } else {
            keysToRemove.forEach(k => cache.remove(k));
        }
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
        const fileId = file.getId();
        if (signatureDataUrl.length < 90000) {
            try {
                CacheService.getScriptCache().put('SIG_B64_' + fileId, signatureDataUrl, 21600);
            } catch (e) { }
        }
        return "drive:" + fileId;
    } catch (err) {
        console.error("Lỗi lưu Drive: " + err.message);
        return signatureDataUrl;
    }
}

function resolveSignatureImageBase64(signVal) {
    if (!signVal) return '';
    if (signVal.startsWith('drive:')) {
        const fileId = signVal.substring(6).trim();
        try {
            const cache = CacheService.getScriptCache();
            const cached = cache.get('SIG_B64_' + fileId);
            if (cached) return cached;
        } catch (e) { }
        try {
            const file = DriveApp.getFileById(fileId);
            const b64 = 'data:image/png;base64,' + Utilities.base64Encode(file.getBlob().getBytes());
            if (b64.length < 90000) {
                try { CacheService.getScriptCache().put('SIG_B64_' + fileId, b64, 21600); } catch (e) { }
            }
            return b64;
        } catch (e) {
            return '';
        }
    }
    return signVal;
}

/**
 * Tải song song hàng loạt chữ ký số từ CacheService & Google Drive REST API (UrlFetchApp.fetchAll)
 * Giảm thời gian tải 200 chữ ký từ ~60s xuống <1.5s (Tránh nghẽn cổ chai N+1 DriveApp.getFileById)
 */
function resolveSignaturesBatch(signVals) {
    const resultMap = {};
    if (!signVals || signVals.length === 0) return resultMap;

    const driveIds = [];
    const uniqueSignVals = [...new Set(signVals.filter(Boolean))];

    uniqueSignVals.forEach(val => {
        if (val.startsWith('drive:')) {
            const fid = val.substring(6).trim();
            if (fid) driveIds.push(fid);
        } else {
            resultMap[val] = val;
        }
    });

    if (driveIds.length === 0) return resultMap;

    const cache = CacheService.getScriptCache();
    const cacheKeys = driveIds.map(id => 'SIG_B64_' + id);
    let cachedObj = {};
    try {
        cachedObj = cache.getAll(cacheKeys) || {};
    } catch (e) { }

    const missingIds = [];
    driveIds.forEach(id => {
        const cVal = cachedObj['SIG_B64_' + id];
        if (cVal) {
            resultMap['drive:' + id] = cVal;
        } else {
            missingIds.push(id);
        }
    });

    if (missingIds.length === 0) return resultMap;

    try {
        const token = ScriptApp.getOAuthToken();
        const toCache = {};
        const CHUNK_SIZE = 50;

        for (let i = 0; i < missingIds.length; i += CHUNK_SIZE) {
            const chunk = missingIds.slice(i, i + CHUNK_SIZE);
            const requests = chunk.map(fid => ({
                url: 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fid) + '?alt=media',
                method: 'get',
                headers: { 'Authorization': 'Bearer ' + token },
                muteHttpExceptions: true
            }));

            const responses = UrlFetchApp.fetchAll(requests);
            responses.forEach((res, idx) => {
                const fid = chunk[idx];
                if (res.getResponseCode() === 200) {
                    const b64 = 'data:image/png;base64,' + Utilities.base64Encode(res.getContent());
                    resultMap['drive:' + fid] = b64;
                    if (b64.length < 90000) {
                        toCache['SIG_B64_' + fid] = b64;
                    }
                } else {
                    resultMap['drive:' + fid] = resolveSignatureImageBase64('drive:' + fid);
                }
            });
        }

        if (Object.keys(toCache).length > 0) {
            try { cache.putAll(toCache, 21600); } catch (e) { }
        }
    } catch (err) {
        missingIds.forEach(fid => {
            resultMap['drive:' + fid] = resolveSignatureImageBase64('drive:' + fid);
        });
    }

    return resultMap;
}

// --- DATA ACCESS & DASHBOARD APIS ---

function getAdminUsers(adminToken, skipVerify) {
    if (!skipVerify && adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const users = getCachedUsersData();
        return users.length > 1 ? users.slice(1) : [];
    } catch (e) {
        return { error: e.toString() };
    }
}

function getAdminCourses(adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    return getCourseStatistics();
}

function getAdminFilters(adminToken) {
    hydrateMemoryFromScriptCache();
    if (_MEM_CACHE.adminFilters) return _MEM_CACHE.adminFilters;
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('admin_filters');
        if (cached) {
            _MEM_CACHE.adminFilters = JSON.parse(cached);
            return _MEM_CACHE.adminFilters;
        }

        const users = getCachedUsersData().slice(1);
        const getUnique = (idx) => [...new Set(users.map(u => String(u[idx] || '').trim()).filter(Boolean))].sort();

        const result = {
            com: getUnique(2),
            group: getUnique(3),
            part: getUnique(4)
        };
        _MEM_CACHE.adminFilters = result;
        cache.put('admin_filters', JSON.stringify(result), 1800);
        return result;
    } catch (e) {
        return { com: [], group: [], part: [] };
    }
}

function getAdminDashboardData(adminToken, skipVerify) {
    if (!skipVerify && !verifyAdminToken(adminToken)) {
        return { error: "Phiên làm việc quản trị đã hết hạn hoặc không hợp lệ. Vui lòng đăng nhập lại!" };
    }
    const users = getAdminUsers(adminToken, true);
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
        const courseData = getCachedCourses();
        const course = courseData.slice(1).find(r => String(r[0]).trim() === String(courseId).trim());
        return course ? course[4] : '';
    } catch (e) {
        return '';
    }
}

function getAvailableCourses() {
    hydrateMemoryFromScriptCache();
    if (_MEM_CACHE.availableCourses) return _MEM_CACHE.availableCourses;
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('available_courses');
        if (cached) {
            _MEM_CACHE.availableCourses = JSON.parse(cached);
            return _MEM_CACHE.availableCourses;
        }

        const courseData = getCachedCourses().slice(1);
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

        _MEM_CACHE.availableCourses = list;
        cache.put('available_courses', JSON.stringify(list), 1800);
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

        const allUsers = userSheet.getDataRange().getDisplayValues();
        const userMap = {};
        allUsers.slice(1).forEach(u => {
            userMap[String(u[1]).trim()] = { name: u[0], id: u[1], com: u[2], group: u[3], part: u[4], role: u[5] };
        });

        const allCourses = getCachedCourses();
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
            clearDashboardCache();
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

// --- USER ASSIGNMENTS & ULTRA LOW-LATENCY WORKFLOW ---

function getMyCourses(userID) {
    try {
        const assignIdx = getCachedAssignmentsIndex();
        const courseData = getCachedCourses();
        const uTarget = String(userID || '').trim();

        const pendingCids = (assignIdx && assignIdx.pendingByUser && assignIdx.pendingByUser[uTarget]) || [];
        if (pendingCids.length === 0) return [];

        const courseMap = {};
        for (let i = 1; i < courseData.length; i++) {
            courseMap[String(courseData[i][0]).trim()] = courseData[i];
        }

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        return pendingCids.map(rawCid => {
            const cid = String(rawCid).trim();
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
                name: courseInfo[1],
                start: courseInfo[2],
                end: courseInfo[3],
                status: status,
                canSign: canSign,
                content: courseInfo[4] || ''
            };
        }).filter(Boolean);
    } catch (e) {
        return [];
    }
}

function confirmTraining(userID, courseID, sign) {
    if (!sign || sign.length < 500) {
        return { success: false, message: "Dữ liệu chữ ký không hợp lệ hoặc bị trống!" };
    }

    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(10000);
    } catch (e) {
        return { success: false, message: "Hệ thống đang bận xử lý nhiều lượt ký đồng thời. Vui lòng thử lại sau 3 giây!" };
    }

    try {
        const uTarget = String(userID).trim();
        const cTarget = String(courseID).trim();
        const courseData = getCachedCourses();

        const course = courseData.slice(1).find(c => String(c[0]).trim() === cTarget);
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

        const db = getDB();
        const assignSheet = db.getSheetByName('Assignments');
        const aData = assignSheet.getDataRange().getDisplayValues();

        for (let i = 1; i < aData.length; i++) {
            if (String(aData[i][2]).trim() === uTarget && String(aData[i][0]).trim() === cTarget) {
                if (aData[i][7] !== '') {
                    return { success: true, message: "Khóa học này đã được ký xác nhận trước đó!" };
                }

                const signRef = saveSignatureToDrive(cTarget, uTarget, sign);
                assignSheet.getRange(i + 1, 8, 1, 2).setValues([[signRef, new Date()]]);

                // Xóa toàn bộ cache chỉ mục phân bổ & thống kê để đồng bộ tức thời
                clearDashboardCache();

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
    hydrateMemoryFromScriptCache();
    if (_MEM_CACHE.courseStats) return _MEM_CACHE.courseStats;
    try {
        const cache = CacheService.getScriptCache();
        const cached = cache.get('course_stats_cache');
        if (cached) {
            _MEM_CACHE.courseStats = JSON.parse(cached);
            return _MEM_CACHE.courseStats;
        }

        const assignIdx = getCachedAssignmentsIndex();
        const assignMap = (assignIdx && assignIdx.statsByCourse) || {};
        const courses = getCachedCourses().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const result = courses.map(c => {
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

        _MEM_CACHE.courseStats = result;
        const serialized = JSON.stringify(result);
        if (serialized.length < 95000) {
            cache.put('course_stats_cache', serialized, 600);
        }
        return result;
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
                clearDashboardCache();
                return true;
            }
        }
        return false;
    } catch (e) {
        return false;
    }
}

// --- SMART REMINDER HUB & INSTANT DATA ENGINE (ZERO DRIVE TRASH) ---

/**
 * Lấy dữ liệu nhắc nhở tức thời (<300ms) cho 1 khóa học hoặc toàn bộ hệ thống
 * Không tạo file tạm trên Google Drive, hỗ trợ xuất Excel bằng SheetJS phía Client & Copy tin nhắn theo Part
 */
function getReminderData(courseId, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const ss = getDB();
        const courses = getCachedCourses().slice(1);
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        const cFilter = courseId ? String(courseId).trim() : null;
        const targetCourses = {};
        const courseMap = {};

        courses.forEach(c => {
            const cid = String(c[0]).trim();
            const start = parseDateFlexible(c[2]);
            const end = parseDateFlexible(c[3]);
            courseMap[cid] = { id: cid, name: c[1], start: c[2], endStr: c[3], end: end };

            if (cFilter) {
                if (cid === cFilter) targetCourses[cid] = c[1];
            } else if (start && today >= start) {
                targetCourses[cid] = c[1];
            }
        });

        const items = [];
        const partMap = {};

        assignments.forEach(a => {
            const cID = String(a[0]).trim();
            if (!targetCourses[cID]) return;

            const part = String(a[6] || 'Chưa phân loại').trim() || 'Chưa phân loại';
            const group = String(a[5] || '-').trim();
            if (!partMap[part]) {
                partMap[part] = { part: part, group: group, total: 0, completed: 0, pending: 0 };
            }
            partMap[part].total++;

            const isCompleted = (a[7] !== '' || a[9] !== '');
            if (isCompleted) {
                partMap[part].completed++;
            } else {
                partMap[part].pending++;
                const cMeta = courseMap[cID] || {};
                const status = (cMeta.end && today > cMeta.end) ? "Quá hạn" : "Chưa hoàn thành";
                items.push({
                    cID: cID,
                    courseName: targetCourses[cID],
                    start: cMeta.start || '',
                    end: cMeta.endStr || '',
                    status: status,
                    uID: String(a[2] || '').trim(),
                    name: String(a[3] || '').trim(),
                    com: String(a[4] || '').trim(),
                    group: group,
                    part: part
                });
            }
        });

        const partSummary = Object.values(partMap)
            .filter(p => p.pending > 0)
            .map(p => ({
                part: p.part,
                group: p.group,
                total: p.total,
                completed: p.completed,
                pending: p.pending,
                rate: p.total > 0 ? Math.round((p.completed / p.total) * 100) : 0
            }))
            .sort((a, b) => b.pending - a.pending);

        let webAppUrl = '';
        try { webAppUrl = ScriptApp.getService().getUrl() || ''; } catch (e) { }

        const scopeTitle = cFilter && courseMap[cFilter]
            ? courseMap[cFilter].name
            : "Tất cả khóa đào tạo đang mở / quá hạn";

        return {
            courseId: cFilter,
            scopeTitle: scopeTitle,
            webAppUrl: webAppUrl,
            totalPending: items.length,
            partSummary: partSummary,
            items: items
        };
    } catch (e) {
        return { error: "Lỗi truy xuất dữ liệu nhắc nhở: " + e.message };
    }
}

/**
 * Gửi email nhắc nhở chủ động ngay lập tức theo yêu cầu của Admin (1-Click Manual Dispatch)
 */
function triggerManualReminderEmail(courseId, adminToken) {
    if (!verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const label = "Đôn đốc tức thời (" + Utilities.formatDate(new Date(), "GMT+7", "HH:mm dd/MM/yyyy") + ")";
        const sentCount = sendIncompleteReport(label, courseId || null);
        if (sentCount === 0) {
            return { error: "Hiện không có học viên nào chưa hoàn thành để gửi email nhắc nhở!" };
        }
        return { success: true, count: sentCount };
    } catch (e) {
        return { error: "Lỗi gửi email nhắc nhở: " + e.message };
    }
}

function exportConsolidatedReminder(adminToken) {
    return getReminderData(null, adminToken);
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

function sendIncompleteReport(timeLabel, filterCourseId) {
    try {
        const ss = getDB();
        const assignments = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);
        const courses = getCachedCourses().slice(1);

        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const cFilter = filterCourseId ? String(filterCourseId).trim() : null;

        const courseMap = {};
        courses.forEach(c => {
            const cid = String(c[0]).trim();
            const start = parseDateFlexible(c[2]);
            const end = parseDateFlexible(c[3]);
            courseMap[cid] = { name: c[1], start: start, end: end, endStr: c[3] };
        });

        const incompleteList = [];
        const partStats = {};

        assignments.forEach(a => {
            const cid = String(a[0]).trim();
            if (cFilter && cid !== cFilter) return;

            const cInfo = courseMap[cid];
            if (!cInfo) return;
            if (!cFilter && cInfo.start && today < cInfo.start) return;

            const part = String(a[6] || 'Chưa phân loại').trim() || 'Chưa phân loại';
            if (!partStats[part]) partStats[part] = { total: 0, completed: 0, pending: 0 };
            partStats[part].total++;

            if (a[7] !== '' || a[9] !== '') {
                partStats[part].completed++;
            } else {
                partStats[part].pending++;
                const status = (cInfo.end && today > cInfo.end) ? "Quá hạn" : "Chưa hoàn thành";
                incompleteList.push({
                    cName: cInfo.name,
                    endStr: cInfo.endStr || '',
                    status: status,
                    uName: a[3],
                    uID: a[2],
                    uGroup: a[5] || '',
                    uPart: part
                });
            }
        });

        if (incompleteList.length === 0) return 0;

        let partRowsHtml = '';
        Object.keys(partStats)
            .filter(p => partStats[p].pending > 0)
            .sort((a, b) => partStats[b].pending - partStats[a].pending)
            .forEach(p => {
                const st = partStats[p];
                const rate = st.total > 0 ? Math.round((st.completed / st.total) * 100) : 0;
                partRowsHtml += `
                <tr>
                    <td style="font-weight:bold;">${p}</td>
                    <td style="text-align:center;">${st.total}</td>
                    <td style="text-align:center;color:#10B981;font-weight:bold;">${st.completed}</td>
                    <td style="text-align:center;color:#EF4444;font-weight:bold;">${st.pending}</td>
                    <td style="text-align:center;font-weight:bold;">${rate}%</td>
                </tr>`;
            });

        let html = `
        <div style="font-family: Arial, sans-serif; color: #1E293B; max-width: 760px; margin: auto; border: 1px solid #CBD5E1; border-radius: 10px; padding: 24px;">
            <div style="border-bottom: 2px solid #2563EB; padding-bottom: 12px; margin-bottom: 16px;">
                <h2 style="color: #0F172A; margin: 0; font-size: 18px;">IQC NEXUS | BÁO CÁO ĐÔN ĐỐC ĐÀO TẠO - ${timeLabel}</h2>
                <div style="font-size: 12px; color: #64748B; margin-top: 4px;">Silence the ads. Let quality make the noise.</div>
            </div>
            <p style="font-size: 14px;">Kính gửi Anh/Chị Quản lý và các Part Leader,</p>
            <p style="font-size: 14px;">Hệ thống ghi nhận tổng số <b style="color:#EF4444;">${incompleteList.length}</b> lượt nhân sự chưa hoàn thành ký xác nhận đào tạo:</p>

            <h4 style="margin: 18px 0 8px 0; color: #1E293B; font-size: 14px;">1. Thống kê tồn đọng theo Bộ phận (Part):</h4>
            <table border="1" cellpadding="7" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 13px; border-color: #CBD5E1; margin-bottom: 18px;">
                <tr style="background-color: #EFF6FF; color: #1E3A8A;">
                    <th style="text-align:left;">Bộ phận (Part)</th>
                    <th>Tổng phân bổ</th>
                    <th>Đã ký</th>
                    <th>Chưa ký</th>
                    <th>Tiến độ</th>
                </tr>
                ${partRowsHtml}
            </table>

            <h4 style="margin: 18px 0 8px 0; color: #1E293B; font-size: 14px;">2. Danh sách nhân sự chưa hoàn thành chi tiết:</h4>
            <table border="1" cellpadding="7" cellspacing="0" style="border-collapse: collapse; width: 100%; font-size: 12.5px; border-color: #CBD5E1;">
                <tr style="background-color: #F8FAFC; text-align: left;">
                    <th>Khóa học</th><th>Họ và tên</th><th>Mã NV</th><th>Bộ phận</th><th>Hạn ký</th><th>Trạng thái</th>
                </tr>`;

        incompleteList.forEach(item => {
            const stColor = item.status === 'Quá hạn' ? '#DC2626' : '#D97706';
            html += `<tr>
                <td>${item.cName}</td>
                <td><b>${item.uName}</b></td>
                <td style="font-family:monospace;">${item.uID}</td>
                <td>${item.uPart}</td>
                <td>${item.endStr}</td>
                <td style="color:${stColor};font-weight:bold;">${item.status}</td>
            </tr>`;
        });

        html += `
            </table>
            <p style="margin-top: 18px; font-size: 12px; color: #64748B; border-top: 1px solid #E2E8F0; padding-top: 10px;">
                <i>Báo cáo tự động từ IQC Nexus • Bản quyền thuộc về Nguyễn Hải Đường (hai.duong). File danh sách đính kèm bên dưới mở trực tiếp bằng Microsoft Excel.</i>
            </p>
        </div>`;

        // Tạo file CSV chuẩn UTF-8 BOM (mở trực tiếp bằng Excel đúng tiếng Việt 100%, 0s tạo file rác trên Drive)
        let csvContent = "\uFEFFKhóa đào tạo,Họ và tên,Mã nhân viên,Group,Bộ phận (Part),Hạn hoàn thành,Trạng thái\r\n";
        incompleteList.forEach(item => {
            const esc = (v) => '"' + String(v || '').replace(/"/g, '""') + '"';
            csvContent += [
                esc(item.cName),
                esc(item.uName),
                esc("'" + item.uID),
                esc(item.uGroup),
                esc(item.uPart),
                esc(item.endStr),
                esc(item.status)
            ].join(',') + "\r\n";
        });

        const attachmentBlob = Utilities.newBlob(csvContent, "text/csv;charset=utf-8", `IQC_NhacNho_${timeLabel.replace(/[\s\(\):\/]/g, '_')}.csv`);

        const recipients = getProp('RECIPIENT_EMAILS', RECIPIENT_EMAILS_DEFAULT);
        MailApp.sendEmail({
            to: recipients,
            subject: `[IQC Nexus] Báo cáo đôn đốc đào tạo (${incompleteList.length} chưa ký) - ${timeLabel}`,
            htmlBody: html,
            attachments: [attachmentBlob]
        });

        return incompleteList.length;
    } catch (e) {
        console.error("Lỗi gửi báo cáo định kỳ: " + e.message);
        return 0;
    }
}

function checkCourseCompletion() {
    try {
        const ss = getDB();
        const courses = getCachedCourses().slice(1);
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

// --- ENGINE BÁO CÁO PDF CHUẨN AUDIT SAMSUNG IQC (BATCH PARALLEL SIGNATURE FETCH) ---

/**
 * Xây dựng nội dung HTML chuẩn Audit ISO/Samsung IQC cho khóa đào tạo
 * Sử dụng resolveSignaturesBatch() để tải song song toàn bộ chữ ký trong <1.5 giây
 */
function buildCourseAuditReportPayload(courseID) {
    const cTarget = String(courseID).trim();
    const ss = getDB();
    const courseData = getCachedCourses().slice(1);
    const assignData = ss.getSheetByName('Assignments').getDataRange().getDisplayValues().slice(1);

    const course = courseData.find(c => String(c[0]).trim() === cTarget);
    if (!course) return { error: "Không tìm thấy khóa học!" };

    const cName = String(course[1]).toUpperCase();
    const cTime = `${course[2]} đến ${course[3]}`;
    const cContent = course[4] || '';

    const allAssigns = assignData.filter(a => String(a[0]).trim() === cTarget);
    if (allAssigns.length === 0) {
        return { error: "Khóa học này chưa có học viên nào được phân bổ!" };
    }

    const completedList = allAssigns.filter(a => a[7] !== '' || a[9] !== '');
    if (completedList.length === 0) {
        return { error: "Hiện chưa có học viên nào hoàn thành hoặc ghi nhận ngoại lệ để xuất báo cáo PDF!" };
    }

    const totalAssigned = allAssigns.length;
    const signedCount = allAssigns.filter(a => a[7] !== '').length;
    const exceptionCount = allAssigns.filter(a => a[7] === '' && a[9] !== '').length;
    const pendingCount = totalAssigned - signedCount - exceptionCount;
    const completionRate = Math.round(((signedCount + exceptionCount) / totalAssigned) * 100);

    // Thống kê theo từng Part
    const partStatsMap = {};
    allAssigns.forEach(a => {
        const part = String(a[6] || 'Chưa phân loại').trim() || 'Chưa phân loại';
        if (!partStatsMap[part]) partStatsMap[part] = { total: 0, signed: 0, exc: 0, pending: 0 };
        partStatsMap[part].total++;
        if (a[7] !== '') partStatsMap[part].signed++;
        else if (a[9] !== '') partStatsMap[part].exc++;
        else partStatsMap[part].pending++;
    });

    // Tải song song toàn bộ chữ ký của khóa học trong 1 lần gọi Batch
    const signRefs = completedList.map(p => p[7]).filter(Boolean);
    const signBase64Map = resolveSignaturesBatch(signRefs);

    const nowStr = Utilities.formatDate(new Date(), "GMT+7", "HH:mm - dd/MM/yyyy");

    let partSummaryHtml = '';
    Object.keys(partStatsMap).sort().forEach(part => {
        const ps = partStatsMap[part];
        const pRate = ps.total > 0 ? Math.round(((ps.signed + ps.exc) / ps.total) * 100) : 0;
        partSummaryHtml += `
            <tr>
                <td class="left"><b>${part}</b></td>
                <td>${ps.total}</td>
                <td style="color:#059669;font-weight:bold;">${ps.signed}</td>
                <td style="color:#d97706;">${ps.exc}</td>
                <td style="color:${ps.pending > 0 ? '#dc2626' : '#64748b'};font-weight:bold;">${ps.pending}</td>
                <td><b>${pRate}%</b></td>
            </tr>`;
    });

    let rowsHtml = '';
    completedList.forEach((p, index) => {
        const name = p[3];
        const uid = p[2];
        const com = p[4];
        const grp = p[5];
        const prt = p[6];
        const rawSign = p[7];
        const signImg = rawSign ? (signBase64Map[rawSign] || '') : '';
        const time = p[8] ? new Date(p[8]).toLocaleString('vi-VN') : '';
        const exc = p[9];

        let signCell = '';
        let metaCell = '';

        if (exc !== '' && rawSign === '') {
            signCell = '<span style="color:#d97706;font-style:italic;font-weight:bold;">Ngoại lệ / Vắng có phép</span>';
            metaCell = exc;
        } else {
            if (signImg && signImg.startsWith('data:image')) {
                signCell = `<img src="${signImg}" class="sign" />`;
            } else {
                signCell = '<span style="color:#059669;font-weight:bold;">Đã ký điện tử</span>';
            }
            metaCell = time;
        }

        rowsHtml += `
            <tr class="sig-row">
                <td>${index + 1}</td>
                <td class="left"><b>${name}</b></td>
                <td style="font-family:monospace;">${uid}</td>
                <td>${com}</td>
                <td>${grp}</td>
                <td>${prt}</td>
                <td>${signCell}</td>
                <td style="font-size:9pt;">${metaCell}</td>
            </tr>`;
    });

    const html = `<!DOCTYPE html>
    <html lang="vi">
    <head>
        <meta charset="utf-8">
        <title>Báo cáo Đào tạo IQC - ${cName}</title>
        <style>
            @page { size: A4 portrait; margin: 14mm 12mm 14mm 12mm; }
            body { font-family: 'Times New Roman', Times, serif; padding: 12px; font-size: 10.5pt; color: #0f172a; line-height: 1.4; }
            .doc-header { width: 100%; border: none; border-bottom: 2px solid #0c65ff; padding-bottom: 8px; margin-bottom: 14px; }
            .doc-header td { border: none; padding: 0; vertical-align: middle; }
            h1 { font-size: 15pt; text-align: center; margin: 6px 0 4px 0; text-transform: uppercase; font-weight: bold; color: #090d16; }
            h2 { font-size: 12.5pt; text-align: center; margin: 0 0 6px 0; color: #1e3a8a; }
            .meta { text-align: center; font-size: 10pt; margin-bottom: 12px; color: #334155; }
            .kpi-table { width: 100%; border-collapse: collapse; margin-bottom: 12px; font-size: 9.5pt; }
            .kpi-table th { background-color: #eff6ff; color: #1e3a8a; border: 1px solid #94a3b8; padding: 5px; text-align: center; }
            .kpi-table td { border: 1px solid #94a3b8; padding: 6px; text-align: center; font-weight: bold; font-size: 10.5pt; }
            .content-box { margin-bottom: 14px; padding: 10px 12px; border: 1px solid #94a3b8; background-color: #f8fafc; font-size: 10pt; border-radius: 4px; }
            .section-title { font-size: 10.5pt; font-weight: bold; text-transform: uppercase; margin: 12px 0 6px 0; color: #0f172a; }
            table.data-grid { width: 100%; border-collapse: collapse; font-size: 9.5pt; margin-bottom: 14px; }
            table.data-grid thead { display: table-header-group; }
            table.data-grid tr.sig-row { page-break-inside: avoid; break-inside: avoid; }
            table.data-grid th, table.data-grid td { border: 1px solid #334155; padding: 5px 6px; vertical-align: middle; text-align: center; }
            table.data-grid th { font-weight: bold; background-color: #f1f5f9; color: #0f172a; }
            .left { text-align: left !important; }
            img.sign { max-height: 36px; max-width: 95px; object-fit: contain; display: block; margin: 0 auto; }
            .approval-table { width: 100%; border-collapse: collapse; margin-top: 18px; page-break-inside: avoid; break-inside: avoid; }
            .approval-table td { border: 1px solid #64748b; width: 33.33%; text-align: center; padding: 8px 6px 52px 6px; vertical-align: top; font-size: 10pt; }
            .doc-footer { margin-top: 18px; text-align: center; font-size: 8.5pt; color: #475569; border-top: 1px solid #cbd5e1; padding-top: 6px; }
            @media print {
                body { padding: 0; }
                .no-print { display: none !important; }
            }
        </style>
    </head>
    <body>
        <table class="doc-header">
            <tr>
                <td style="text-align: left;">
                    <div style="font-size: 13pt; font-weight: bold; color: #0c65ff; letter-spacing: 0.02em;">IQC NEXUS | SAMSUNG QUALITY ENGINEERING</div>
                    <div style="font-size: 8.5pt; color: #475569; font-style: italic;">Silence the ads. Let quality make the noise.</div>
                </td>
                <td style="text-align: right; font-size: 8.5pt; color: #334155; font-family: monospace;">
                    <div><b>MÃ KHÓA:</b> ${cTarget}</div>
                    <div><b>XUẤT LÚC:</b> ${nowStr}</div>
                </td>
            </tr>
        </table>

        <h1>BIÊN BẢN XÁC NHẬN HOÀN THÀNH ĐÀO TẠO</h1>
        <h2>KHÓA ĐÀO TẠO: ${cName}</h2>
        <div class="meta">Thời hạn hiệu lực: <b>${cTime}</b></div>

        <table class="kpi-table">
            <tr>
                <th>Tổng quân số phân bổ</th>
                <th>Đã ký xác nhận</th>
                <th>Vắng có phép / Ngoại lệ</th>
                <th>Chưa hoàn thành</th>
                <th>Tỷ lệ hoàn thành</th>
            </tr>
            <tr>
                <td>${totalAssigned}</td>
                <td style="color: #059669;">${signedCount}</td>
                <td style="color: #d97706;">${exceptionCount}</td>
                <td style="color: ${pendingCount > 0 ? '#dc2626' : '#059669'};">${pendingCount}</td>
                <td style="color: #0c65ff;">${completionRate}%</td>
            </tr>
        </table>

        <div class="content-box">
            <div style="font-weight: bold; margin-bottom: 4px; text-transform: uppercase; font-size: 9.5pt; color: #1e3a8a;">I. Nội dung & Tiêu chuẩn Đào tạo:</div>
            <div>${cContent}</div>
        </div>

        <div class="section-title">II. Thống kê tiến độ theo Bộ phận (Part):</div>
        <table class="data-grid">
            <thead>
                <tr>
                    <th class="left" style="width: 28%">Bộ phận (Part)</th>
                    <th style="width: 18%">Tổng nhân sự</th>
                    <th style="width: 18%">Đã ký</th>
                    <th style="width: 18%">Ngoại lệ</th>
                    <th style="width: 18%">Chưa ký</th>
                    <th style="width: 18%">Tỷ lệ đạt</th>
                </tr>
            </thead>
            <tbody>
                ${partSummaryHtml}
            </tbody>
        </table>

        <div class="section-title">III. Danh sách Chữ ký Điện tử & Chứng thực Học viên (${completedList.length}/${totalAssigned}):</div>
        <table class="data-grid">
            <thead>
                <tr>
                    <th style="width: 5%">STT</th>
                    <th style="width: 22%">Họ và tên</th>
                    <th style="width: 11%">Mã NV</th>
                    <th style="width: 9%">Company</th>
                    <th style="width: 10%">Group</th>
                    <th style="width: 10%">Part</th>
                    <th style="width: 17%">Chữ ký điện tử</th>
                    <th style="width: 16%">Thời gian ký / Ghi chú</th>
                </tr>
            </thead>
            <tbody>
                ${rowsHtml}
            </tbody>
        </table>

        <table class="approval-table">
            <tr>
                <td>
                    <b>NGƯỜI LẬP BIỂU</b><br>
                    <span style="font-size: 8.5pt; color: #64748b;">(Ký và ghi rõ họ tên)</span>
                </td>
                <td>
                    <b>XEM XÉT (PART LEADER)</b><br>
                    <span style="font-size: 8.5pt; color: #64748b;">(Ký và ghi rõ họ tên)</span>
                </td>
                <td>
                    <b>PHÊ DUYỆT (IQC MANAGER)</b><br>
                    <span style="font-size: 8.5pt; color: #64748b;">(Ký và ghi rõ họ tên)</span>
                </td>
            </tr>
        </table>

        <div class="doc-footer">
            Silence the ads. Let quality make the noise. • Bản quyền thuộc về Nguyễn Hải Đường (hai.duong) • IQC Quality Engineering Platform
        </div>
    </body>
    </html>`;

    return {
        courseName: cName,
        filename: `BaoCao_IQC_${cName.replace(/[\\/:*?"<>|]/g, '_')}.pdf`,
        html: html
    };
}

/**
 * Trả về HTML báo cáo chuẩn Audit để in Vector trực tiếp hoặc lưu PDF tức thì (1 giây) trên trình duyệt
 */
function getCourseAuditReportHtml(courseID, adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        return buildCourseAuditReportPayload(courseID);
    } catch (e) {
        return { error: "Lỗi kết xuất bản in Audit: " + e.message };
    }
}

/**
 * Kết xuất file .pdf trực tiếp từ Server (đã tối ưu tải chữ ký song song bằng resolveSignaturesBatch)
 */
function exportCourseReport(courseID, adminToken) {
    if (adminToken && !verifyAdminToken(adminToken)) return { error: "Không có quyền quản trị!" };
    try {
        const payload = buildCourseAuditReportPayload(courseID);
        if (payload.error) return payload;

        const blob = Utilities.newBlob(payload.html, MimeType.HTML, "audit_report.html");
        const pdfBlob = blob.getAs(MimeType.PDF).setName(payload.filename);
        const finalBase64 = Utilities.base64Encode(pdfBlob.getBytes());

        return { filename: payload.filename, base64: finalBase64 };
    } catch (e) {
        return { error: "Lỗi tạo báo cáo PDF: " + e.message };
    }
}
