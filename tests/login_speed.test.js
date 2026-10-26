/**
 * SMART-TEST SUITE: Login Speed & Multi-Tier Cache Integrity
 * @blast-radius: core, db, api
 */
const fs = require('fs');
const path = require('path');
const vm = require('vm');
const assert = require('assert');

function createGasSandbox(simulatedSheetLatencyMs = 600) {
    const metrics = {
        openByIdCalls: 0,
        sheetReadCalls: 0,
        propCalls: 0,
        cacheGets: 0,
        cachePuts: 0,
        simulatedLatencyMs: 0
    };

    const cacheStore = new Map();
    const sheetsData = {
        Users: [
            ['Name', 'UserID', 'Company', 'Group', 'Part', 'Role', 'Status'],
            ['Nguyễn Hải Đường', '10545998', 'SEV', 'IQC G', 'IQC 1P', 'user', 'Active'],
            ['Trần Văn B', '10545999', 'SEV', 'IQC G', 'IQC 2P', 'user', 'Inactive'],
            ['Lê Thị Mới', '10546000', 'SEV', 'IQC G', 'IQC 1P', 'user', 'New-Reg']
        ],
        Courses: [
            ['ID', 'Name', 'Start', 'End', 'Content', 'Status', 'SignEnabled'],
            ['TRN-001', 'Đào tạo Tiêu chuẩn Ngoại quan', '01/01/2026', '31/12/2026', '<p>Nội dung chi tiết rất dài...</p>'.repeat(50), 'Opening', 'true'],
            ['TRN-002', 'Đào tạo An toàn ESD', '01/01/2026', '31/12/2026', '<p>Tài liệu ESD...</p>'.repeat(50), 'Opening', 'true']
        ],
        Assignments: [
            ['CourseID', 'CourseName', 'UserID', 'UserName', 'Com', 'Group', 'Part', 'Sign', 'Time', 'Exception'],
            ['TRN-001', 'Đào tạo Tiêu chuẩn Ngoại quan', '10545998', 'Nguyễn Hải Đường', 'SEV', 'IQC G', 'IQC 1P', '', '', ''],
            ['TRN-002', 'Đào tạo An toàn ESD', '10545998', 'Nguyễn Hải Đường', 'SEV', 'IQC G', 'IQC 1P', 'drive:sig123', '2026-10-20', '']
        ]
    };

    function makeSheet(name) {
        return {
            getDataRange() {
                return {
                    getDisplayValues() {
                        metrics.sheetReadCalls++;
                        metrics.simulatedLatencyMs += simulatedSheetLatencyMs;
                        return sheetsData[name].map(r => [...r]);
                    },
                    getValues() {
                        metrics.sheetReadCalls++;
                        metrics.simulatedLatencyMs += simulatedSheetLatencyMs;
                        return sheetsData[name].map(r => [...r]);
                    }
                };
            },
            getRange(row, col, numRows = 1, numCols = 1) {
                return {
                    setNumberFormat() { return this; },
                    setValue(val) {
                        sheetsData[name][row - 1][col - 1] = String(val);
                    },
                    setValues(vals) {
                        for (let r = 0; r < vals.length; r++) {
                            if (!sheetsData[name][row - 1 + r]) sheetsData[name][row - 1 + r] = [];
                            for (let c = 0; c < vals[r].length; c++) {
                                sheetsData[name][row - 1 + r][col - 1 + c] = String(vals[r][c]);
                            }
                        }
                    },
                    clearContent() { }
                };
            },
            getLastRow() {
                return sheetsData[name].length;
            }
        };
    }

    const sandbox = {
        console,
        Date,
        Math,
        JSON,
        Set,
        Map,
        Array,
        Object,
        String,
        Number,
        RegExp,
        parseInt,
        isNaN,
        PropertiesService: {
            getScriptProperties() {
                return {
                    getProperty(key) {
                        metrics.propCalls++;
                        metrics.simulatedLatencyMs += 80;
                        return null;
                    },
                    getProperties() {
                        metrics.propCalls++;
                        metrics.simulatedLatencyMs += 80;
                        return {};
                    }
                };
            }
        },
        SpreadsheetApp: {
            openById(id) {
                metrics.openByIdCalls++;
                metrics.simulatedLatencyMs += 450;
                return {
                    getSheetByName(name) {
                        return sheetsData[name] ? makeSheet(name) : null;
                    }
                };
            }
        },
        CacheService: {
            getScriptCache() {
                return {
                    get(key) {
                        metrics.cacheGets++;
                        metrics.simulatedLatencyMs += 30;
                        return cacheStore.has(key) ? cacheStore.get(key) : null;
                    },
                    getAll(keys) {
                        metrics.cacheGets++;
                        metrics.simulatedLatencyMs += 40;
                        const out = {};
                        keys.forEach(k => { if (cacheStore.has(k)) out[k] = cacheStore.get(k); });
                        return out;
                    },
                    put(key, val) {
                        metrics.cachePuts++;
                        metrics.simulatedLatencyMs += 30;
                        cacheStore.set(key, String(val));
                    },
                    putAll(obj) {
                        metrics.cachePuts++;
                        metrics.simulatedLatencyMs += 40;
                        Object.keys(obj).forEach(k => cacheStore.set(k, String(obj[k])));
                    },
                    remove(key) {
                        cacheStore.delete(key);
                    },
                    removeAll(keys) {
                        keys.forEach(k => cacheStore.delete(k));
                    }
                };
            }
        },
        LockService: {
            getScriptLock() {
                return {
                    waitLock() { },
                    releaseLock() { }
                };
            }
        },
        DriveApp: {
            getRootFolder() {
                return {
                    getFoldersByName() { return { hasNext: () => false }; },
                    createFolder() { return { getFoldersByName: () => ({ hasNext: () => false }), createFile: () => ({ getId: () => 'fid_1' }) }; }
                };
            },
            getFileById() {
                return {
                    getParents: () => ({ hasNext: () => false })
                };
            }
        },
        Utilities: {
            getUuid: () => 'uuid-test-1234',
            base64Decode: () => Buffer.from('test'),
            base64Encode: () => 'dGVzdA==',
            newBlob: () => ({ getId: () => 'blob1' }),
            formatDate: () => '10:00 27/10/2026'
        },
        MimeType: { PNG: 'image/png', HTML: 'text/html', PDF: 'application/pdf' }
    };

    const code = fs.readFileSync(path.join(__dirname, '..', 'Code.js'), 'utf8');
    vm.createContext(sandbox);
    vm.runInContext(code, sandbox);

    return {
        sandbox,
        metrics,
        resetMetrics() {
            metrics.openByIdCalls = 0;
            metrics.sheetReadCalls = 0;
            metrics.propCalls = 0;
            metrics.cacheGets = 0;
            metrics.cachePuts = 0;
            metrics.simulatedLatencyMs = 0;
        }
    };
}

function runTests() {
    console.log('=== RUNNING SMART-TEST SUITE (@blast-radius: core, db, api) ===');
    const env = createGasSandbox(600);

    // 1. Measure Cold Admin Login
    env.resetMetrics();
    const adminRes = env.sandbox.checkLogin('admin', '@nhd16488');
    assert.strictEqual(adminRes.role, 'admin', 'Admin login must return role=admin');
    console.log(`[Metric] Cold Admin Login -> openById: ${env.metrics.openByIdCalls}, sheetReads: ${env.metrics.sheetReadCalls}, simulatedLatency: ${env.metrics.simulatedLatencyMs}ms`);

    // 2. Measure Warm User Login (After Prewarm)
    if (typeof env.sandbox.prewarmSystem === 'function') {
        env.sandbox.prewarmSystem();
    }
    env.resetMetrics();
    const userRes = env.sandbox.checkLogin('10545998', '@iqc6368');
    assert.strictEqual(userRes.role, 'user', 'Active user login must succeed');
    assert.strictEqual(userRes.courses.length, 1, 'User 10545998 should have 1 unsigned course (TRN-001)');
    assert.strictEqual(userRes.courses[0].id, 'TRN-001');
    console.log(`[Metric] Warm User Login  -> openById: ${env.metrics.openByIdCalls}, sheetReads: ${env.metrics.sheetReadCalls}, simulatedLatency: ${env.metrics.simulatedLatencyMs}ms`);

    // 3. Measure Warm Admin Login (After Prewarm)
    env.resetMetrics();
    const warmAdminRes = env.sandbox.checkLogin('admin', '@nhd16488');
    assert.strictEqual(warmAdminRes.role, 'admin');
    console.log(`[Metric] Warm Admin Login -> openById: ${env.metrics.openByIdCalls}, sheetReads: ${env.metrics.sheetReadCalls}, simulatedLatency: ${env.metrics.simulatedLatencyMs}ms`);

    // 4. Verify Inactive User Rejection
    const inactiveRes = env.sandbox.checkLogin('10545999', '@iqc6368');
    assert.ok(inactiveRes.error && inactiveRes.error.includes('khóa'), 'Inactive user must be rejected');

    // 5. Assertions on Performance SLA (@blast-radius: db, api)
    assert.strictEqual(env.metrics.sheetReadCalls, 0, `FAIL: Warm login must perform 0 Spreadsheet reads, but performed ${env.metrics.sheetReadCalls}`);
    assert.ok(env.metrics.simulatedLatencyMs <= 150, `FAIL: Warm login simulated RPC latency must be <= 150ms, but was ${env.metrics.simulatedLatencyMs}ms`);

    // 6. Verify Cache Invalidation on Signature Confirmation
    const fakeSign = 'data:image/png;base64,' + 'A'.repeat(600);
    const signRes = env.sandbox.confirmTraining('10545998', 'TRN-001', fakeSign);
    assert.strictEqual(signRes.success, true, 'confirmTraining must succeed');

    const afterSignLogin = env.sandbox.checkLogin('10545998', '@iqc6368');
    assert.strictEqual(afterSignLogin.courses.length, 0, 'After signing TRN-001, user should have 0 pending courses');

    // 7. Verify Client-Side SWR & Speculative Login Engine (@blast-radius: ui)
    const indexHtml = fs.readFileSync(path.join(__dirname, '..', 'Index.html'), 'utf8');
    assert.ok(indexHtml.includes('prewarmSystem()'), 'Index.html must trigger background prewarmSystem() on load');
    assert.ok(indexHtml.includes('getOrStartLoginPromise'), 'Index.html must include speculative in-flight login prefetch');
    assert.ok(indexHtml.includes('nexus_swr_'), 'Index.html must include <15ms Stale-While-Revalidate instant auth cache');

    console.log('✅ ALL SMART-TEST ASSERTIONS PASSED!');
}

runTests();
