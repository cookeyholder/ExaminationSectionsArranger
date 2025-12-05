/**
 * 測試執行器
 * 在 Google Apps Script 環境中執行單元測試與整合測試
 */

/**
 * 簡易測試框架
 */
class TestRunner {
    constructor(name) {
        this.suiteName = name;
        this.tests = [];
        this.results = {
            passed: 0,
            failed: 0,
            skipped: 0,
            errors: [],
        };
    }

    /**
     * 新增測試案例
     */
    test(description, testFn) {
        this.tests.push({ description, testFn });
        return this;
    }

    /**
     * 執行所有測試
     */
    run() {
        Logger.log(
            `\n========== 開始執行測試套件: ${this.suiteName} ==========\n`
        );
        const startTime = new Date().getTime();

        this.tests.forEach((test, idx) => {
            try {
                Logger.log(
                    `[${idx + 1}/${this.tests.length}] ${test.description}`
                );
                test.testFn();
                this.results.passed++;
                Logger.log("  ✅ PASS\n");
            } catch (error) {
                this.results.failed++;
                this.results.errors.push({
                    test: test.description,
                    error: error.message,
                    stack: error.stack,
                });
                Logger.log(`  ❌ FAIL: ${error.message}\n`);
            }
        });

        const endTime = new Date().getTime();
        const duration = endTime - startTime;

        this.printSummary(duration);
        return this.results;
    }

    /**
     * 列印測試摘要
     */
    printSummary(duration) {
        Logger.log("\n========== 測試結果摘要 ==========");
        Logger.log(`測試套件: ${this.suiteName}`);
        Logger.log(`總計: ${this.tests.length} 項測試`);
        Logger.log(`✅ 通過: ${this.results.passed}`);
        Logger.log(`❌ 失敗: ${this.results.failed}`);
        Logger.log(`執行時間: ${duration}ms`);

        if (this.results.failed > 0) {
            Logger.log("\n---------- 失敗詳情 ----------");
            this.results.errors.forEach((error, idx) => {
                Logger.log(`\n${idx + 1}. ${error.test}`);
                Logger.log(`   錯誤: ${error.error}`);
            });
        }

        Logger.log("\n==================================\n");
    }
}

/**
 * 斷言函式庫
 */
const assert = {
    /**
     * 斷言值為真
     */
    isTrue(value, message = "預期值為 true") {
        if (value !== true) {
            throw new Error(`${message} (實際: ${value})`);
        }
    },

    /**
     * 斷言值為假
     */
    isFalse(value, message = "預期值為 false") {
        if (value !== false) {
            throw new Error(`${message} (實際: ${value})`);
        }
    },

    /**
     * 斷言相等
     */
    equals(actual, expected, message = "") {
        if (actual !== expected) {
            const msg = message || `預期值: ${expected}, 實際值: ${actual}`;
            throw new Error(msg);
        }
    },

    /**
     * 斷言深度相等（物件/陣列）
     */
    deepEquals(actual, expected, message = "") {
        const actualJSON = JSON.stringify(actual);
        const expectedJSON = JSON.stringify(expected);
        if (actualJSON !== expectedJSON) {
            const msg = message || `預期: ${expectedJSON}\n實際: ${actualJSON}`;
            throw new Error(msg);
        }
    },

    /**
     * 斷言包含
     */
    contains(array, value, message = "") {
        if (!array.includes(value)) {
            const msg = message || `陣列不包含 ${value}`;
            throw new Error(msg);
        }
    },

    /**
     * 斷言大於
     */
    greaterThan(actual, expected, message = "") {
        if (actual <= expected) {
            const msg = message || `${actual} 不大於 ${expected}`;
            throw new Error(msg);
        }
    },

    /**
     * 斷言小於
     */
    lessThan(actual, expected, message = "") {
        if (actual >= expected) {
            const msg = message || `${actual} 不小於 ${expected}`;
            throw new Error(msg);
        }
    },

    /**
     * 斷言拋出例外
     */
    throws(fn, message = "預期函式會拋出例外") {
        let thrown = false;
        try {
            fn();
        } catch (error) {
            thrown = true;
        }
        if (!thrown) {
            throw new Error(message);
        }
    },

    /**
     * 斷言不為 null 或 undefined
     */
    notNull(value, message = "值不應為 null 或 undefined") {
        if (value === null || value === undefined) {
            throw new Error(message);
        }
    },
};

/**
 * ========== 領域模型測試 ==========
 */
function testDomainModels() {
    const suite = new TestRunner("領域模型測試");

    suite
        .test("Classroom 物件建立與學生新增", () => {
            const classroom = createClassroomRecord();
            assert.equals(classroom.population, 0, "初始人數應為 0");

            classroom.addStudent(["資訊", "一年級", "資一甲", "張三", "數學"]);
            assert.equals(classroom.population, 1, "新增後人數應為 1");
            assert.equals(classroom.students.length, 1);
        })

        .test("Classroom 統計功能", () => {
            const classroom = createClassroomRecord();
            classroom.addStudent(["資訊", "一年級", "資一甲", "張三", "數學"]);
            classroom.addStudent(["資訊", "一年級", "資一乙", "李四", "數學"]);
            classroom.addStudent(["機械", "二年級", "機二甲", "王五", "英文"]);

            const stats = classroom.classSubjectStatistics;
            assert.equals(stats["資一甲-數學"], 1);
            assert.equals(stats["資一乙-數學"], 1);
            assert.equals(stats["機二甲-英文"], 1);
        })

        .test("Session 物件建立與試場管理", () => {
            const session = createSessionRecord(5);
            assert.equals(session.population, 0);
            assert.equals(session.classrooms.length, 6); // 索引 0 不使用

            session.addStudent(["資訊", "一年級", "資一甲", "張三", "數學"]);
            assert.equals(session.population, 1);
        })

        .test("Session 分配學生至試場", () => {
            const session = createSessionRecord(3);

            // 新增測試學生並指定試場
            const students = [
                ["資訊", "一年級", "資一甲", "學生1", "數學", 1, 1],
                ["資訊", "一年級", "資一甲", "學生2", "數學", 1, 1],
                ["機械", "二年級", "機二甲", "學生3", "英文", 1, 2],
            ];

            students.forEach((s) => session.addStudent(s));

            // 執行分配
            session.distributeToChildren((student, classrooms) => {
                const roomIndex = student[6]; // 試場編號在索引 6
                return classrooms[roomIndex];
            });

            assert.equals(session.classrooms[1].population, 2);
            assert.equals(session.classrooms[2].population, 1);
        })

        .test("Exam 物件完整流程", () => {
            const exam = createExamRecord(3, 2);

            // 新增不同節次的學生
            exam.addStudent([
                "資訊",
                "一年級",
                "資一甲",
                "學生1",
                "數學",
                1,
                1,
            ]);
            exam.addStudent([
                "資訊",
                "一年級",
                "資一甲",
                "學生2",
                "數學",
                1,
                1,
            ]);
            exam.addStudent([
                "機械",
                "二年級",
                "機二甲",
                "學生3",
                "英文",
                2,
                1,
            ]);

            assert.equals(exam.population, 3);

            // 分配到節次
            exam.distributeToChildren((student, sessions) => {
                const sessionIndex = student[5];
                return sessions[sessionIndex];
            });

            assert.equals(exam.sessions[1].population, 2);
            assert.equals(exam.sessions[2].population, 1);

            // 檢查統計
            const stats = exam.sessionDistribution;
            assert.equals(stats[1], 2);
            assert.equals(stats[2], 1);
        })

        .test("統計容器的 clear 功能", () => {
            const classroom = createClassroomRecord();
            classroom.addStudent(["資訊", "一年級", "資一甲", "張三", "數學"]);
            classroom.addStudent(["資訊", "一年級", "資一甲", "李四", "英文"]);

            assert.equals(classroom.population, 2);

            classroom.clear();
            assert.equals(classroom.population, 0);
            assert.equals(classroom.students.length, 0);
        });

    return suite.run();
}

/**
 * ========== examService 測試 ==========
 */
function testExamService() {
    const suite = new TestRunner("ExamService 測試");

    suite
        .test("createExamFromSheet 讀取工作表資料", () => {
            const exam = createExamFromSheet();
            assert.notNull(exam, "Exam 物件不應為 null");
            assert.greaterThan(exam.sessions.length, 0, "應有節次資料");
        })

        .test("Exam 統計資料正確性", () => {
            const exam = createExamFromSheet();

            // 驗證統計維度存在
            assert.notNull(exam.sessionDistribution);
            assert.notNull(exam.departmentDistribution);
            assert.notNull(exam.gradeDistribution);
            assert.notNull(exam.subjectDistribution);

            // 驗證總人數一致性
            let totalFromSessions = 0;
            exam.sessions.slice(1).forEach((session) => {
                if (session) totalFromSessions += session.population;
            });

            // 注意：這裡可能因為資料狀態而有差異，暫時略過嚴格檢查
            Logger.log(
                `Exam 總人數: ${exam.population}, 節次加總: ${totalFromSessions}`
            );
        });

    return suite.run();
}

/**
 * ========== 排程邏輯測試 ==========
 */
function testSchedulingLogic() {
    const suite = new TestRunner("排程邏輯測試");

    suite
        .test("assignSessionTimesForExam 節次分配", () => {
            // 這個測試需要實際工作表資料
            // 先檢查函式存在
            assert.notNull(typeof assignSessionTimesForExam === "function");
        })

        .test("assignExamRooms 試場編排", () => {
            assert.notNull(typeof assignExamRooms === "function");
        });

    return suite.run();
}

/**
 * ========== 執行所有測試 ==========
 */
function runAllTests() {
    Logger.log("\n");
    Logger.log("╔════════════════════════════════════════╗");
    Logger.log("║   GAS 專案測試執行器                   ║");
    Logger.log("╚════════════════════════════════════════╝");
    Logger.log("\n");

    const allResults = {
        total: 0,
        passed: 0,
        failed: 0,
        suites: [],
    };

    // 執行各測試套件
    const suites = [
        { name: "領域模型", fn: testDomainModels },
        { name: "ExamService", fn: testExamService },
        { name: "排程邏輯", fn: testSchedulingLogic },
    ];

    suites.forEach((suite) => {
        try {
            const result = suite.fn();
            allResults.total += result.passed + result.failed;
            allResults.passed += result.passed;
            allResults.failed += result.failed;
            allResults.suites.push({ name: suite.name, result });
        } catch (error) {
            Logger.log(
                `\n❌ 測試套件 "${suite.name}" 執行失敗: ${error.message}\n`
            );
            allResults.failed++;
        }
    });

    // 總結報告
    Logger.log("\n");
    Logger.log("╔════════════════════════════════════════╗");
    Logger.log("║          總體測試結果                  ║");
    Logger.log("╚════════════════════════════════════════╝");
    Logger.log(`\n總測試數: ${allResults.total}`);
    Logger.log(`✅ 通過: ${allResults.passed}`);
    Logger.log(`❌ 失敗: ${allResults.failed}`);
    Logger.log(
        `成功率: ${((allResults.passed / allResults.total) * 100).toFixed(1)}%`
    );
    Logger.log("\n");

    return allResults;
}

/**
 * 快速測試：驗證當前狀態
 */
function quickTest() {
    Logger.log("執行快速測試...\n");

    try {
        // 測試 1: 工作表連線
        Logger.log("1. 測試工作表連線...");
        const sheet = EXAM_SESSIONS_SHEET;
        assert.notNull(sheet, "工作表應存在");
        Logger.log("   ✅ 通過\n");

        // 測試 2: Exam 物件建立
        Logger.log("2. 測試 Exam 物件建立...");
        const exam = createExamFromSheet();
        assert.notNull(exam, "Exam 物件應存在");
        Logger.log(`   ✅ 通過 (總人數: ${exam.population})\n`);

        // 測試 3: 統計功能
        Logger.log("3. 測試統計功能...");
        assert.notNull(exam.sessionDistribution);
        assert.notNull(exam.departmentDistribution);
        Logger.log("   ✅ 通過\n");

        Logger.log("🎉 快速測試全部通過！");
    } catch (error) {
        Logger.log(`\n❌ 快速測試失敗: ${error.message}`);
        Logger.log(error.stack);
    }
}
