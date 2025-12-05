/**
 * Pipeline 處理器模組
 *
 * 此模組提供高效能的資料處理 Pipeline，採用「單次讀取-連續處理-單次寫入」模式。
 * 所有 Pipeline 函式以 `pipeline_` 前綴命名，與原有的單步驟函式區隔。
 *
 * 設計原則：
 * 1. 資料在記憶體中連續傳遞，直到最後才寫回試算表
 * 2. 每個 Pipeline 函式都是純函式，接收資料並回傳處理結果
 * 3. 可組合成不同的處理流程，彈性應對各種使用情境
 *
 * 使用範例：
 * ```javascript
 * // 完整流程
 * const result = pipeline_executeFullScheduling();
 *
 * // 部分流程（手動調整後）
 * const result = pipeline_executePostAdjustment();
 * ```
 */

// ============================================================================
// Pipeline 資料結構
// ============================================================================

/**
 * 建立 Pipeline 上下文物件
 *
 * 此物件在整個 Pipeline 過程中傳遞，包含所有需要的資料和配置。
 *
 * @returns {Object} Pipeline 上下文
 */
function pipeline_createContext() {
    return {
        students: [], // 學生資料陣列
        headerRow: [], // 標題列
        columns: {}, // 欄位索引
        params: {}, // 參數配置
        sessionRules: {}, // 科目節次規則
        sessionTimes: [], // 節次時間資料
        warnings: [], // 警告訊息
        stats: {}, // 統計資訊
    };
}

// ============================================================================
// Pipeline I/O 函式
// ============================================================================

/**
 * 載入所有資料到 Pipeline 上下文
 *
 * 一次性讀取所有需要的資料，建立完整的處理上下文。
 *
 * @returns {Object} 已初始化的 Pipeline 上下文
 */
function pipeline_loadData() {
    const ctx = pipeline_createContext();

    // 一次性讀取所有資料
    const rawData = loadAllData();

    ctx.headerRow = rawData.students[0];
    ctx.students = rawData.students.slice(1);
    ctx.columns = buildColumnIndices(ctx.headerRow);
    ctx.params = parseParameters(rawData.parameters);
    ctx.sessionRules = parseSessionRules(rawData.sessionRules);
    ctx.sessionTimes = rawData.sessionTimes;

    return ctx;
}

/**
 * 將處理結果寫回試算表
 *
 * @param {Object} ctx - Pipeline 上下文
 */
function pipeline_saveData(ctx) {
    saveAllData(ctx.students);
    FILTERED_RESULT_SHEET.getRange("I:J").setNumberFormat("#,##0");
}

// ============================================================================
// Pipeline 核心處理函式（純函式）
// ============================================================================

/**
 * 安排共同科目節次
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_scheduleCommonSubjects(ctx) {
    ctx.students.forEach(function (student) {
        const subjectName = student[ctx.columns.subject];
        const preferredSession = ctx.sessionRules[subjectName];
        if (preferredSession != null) {
            student[ctx.columns.session] = preferredSession;
        }
    });

    return ctx;
}

/**
 * 安排專業科目節次
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_scheduleSpecializedSubjects(ctx) {
    const sessionCapacity = 0.9 * ctx.params.sessionCapacity;
    const maxSessionCount = ctx.params.maxSessionCount;
    const columns = ctx.columns;

    // 只統計尚未分配節次的學生
    const deptGradeSubjectCounts = {};
    ctx.students.forEach(function (student) {
        if (student[columns.session] !== 0 && student[columns.session] !== "") {
            return;
        }
        const key =
            student[columns.department] +
            student[columns.grade] +
            "_" +
            student[columns.subject];
        deptGradeSubjectCounts[key] = (deptGradeSubjectCounts[key] || 0) + 1;
    });

    const sortedCounts = Object.entries(deptGradeSubjectCounts).sort(
        compareCountDescending
    );

    // 節次統計
    const sessionStats = {};
    for (let i = 1; i <= maxSessionCount; i++) {
        sessionStats[i] = { population: 0, deptGrade: {} };
    }

    // 分配學生到節次
    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        const session = sessionStats[sessionNumber];

        for (let i = 0; i < sortedCounts.length; i++) {
            const [deptGradeSubjectKey, studentCount] = sortedCounts[i];
            const deptGradeKey = deptGradeSubjectKey.substring(
                0,
                deptGradeSubjectKey.indexOf("_")
            );

            // 檢查互斥規則
            if (Object.keys(session.deptGrade).includes(deptGradeKey)) {
                continue;
            }

            // 檢查容量
            if (studentCount + session.population > sessionCapacity) {
                continue;
            }

            // 分配學生
            ctx.students.forEach(function (student) {
                const studentKey =
                    student[columns.department] +
                    student[columns.grade] +
                    "_" +
                    student[columns.subject];
                const currentSession = student[columns.session];
                if (
                    studentKey === deptGradeSubjectKey &&
                    (currentSession === 0 || currentSession === "")
                ) {
                    student[columns.session] = sessionNumber;
                    session.population++;
                    session.deptGrade[deptGradeKey] =
                        (session.deptGrade[deptGradeKey] || 0) + 1;
                }
            });
        }
    }

    // 檢查未分配
    const unscheduledCount = ctx.students.filter(function (s) {
        const session = s[columns.session];
        return session === 0 || session === "";
    }).length;

    if (unscheduledCount > 0) {
        ctx.warnings.push(
            "無法將所有人排入 " +
                maxSessionCount +
                " 節，請檢查是否有某科年級須補考過多科目！"
        );
    }

    ctx.stats.unscheduledCount = unscheduledCount;
    return ctx;
}

/**
 * 安排試場
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_assignRooms(ctx) {
    const columns = ctx.columns;
    const maxSessionCount = ctx.params.maxSessionCount;
    const maxRoomCount = ctx.params.maxRoomCount;
    const maxStudentsPerRoom = ctx.params.maxStudentsPerRoom;
    const maxSubjectsPerRoom = ctx.params.maxSubjectsPerRoom;

    // 重置所有試場
    ctx.students.forEach(function (student) {
        student[columns.room] = 0;
    });

    let allScheduled = true;

    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        // 取得本節次學生
        const sessionStudents = ctx.students.filter(function (s) {
            return s[columns.session] === sessionNumber;
        });

        if (sessionStudents.length === 0) continue;

        // 計算班級科目統計
        const deptClassSubjectCounts = {};
        sessionStudents.forEach(function (student) {
            const key = student[columns.class] + student[columns.subject];
            deptClassSubjectCounts[key] =
                (deptClassSubjectCounts[key] || 0) + 1;
        });

        const sortedCounts = Object.entries(deptClassSubjectCounts).sort(
            compareCountDescending
        );

        // 試場統計
        const roomStats = {};
        for (let r = 1; r <= maxRoomCount; r++) {
            roomStats[r] = { population: 0, classSubject: {} };
        }

        for (let roomNumber = 1; roomNumber <= maxRoomCount; roomNumber++) {
            const room = roomStats[roomNumber];
            const scheduledKeys = [];

            for (let i = 0; i < sortedCounts.length; i++) {
                const [classSubjectKey, count] = sortedCounts[i];

                if (scheduledKeys.includes(classSubjectKey)) continue;
                if (count + room.population > maxStudentsPerRoom) continue;
                if (
                    Object.keys(room.classSubject).length + 1 >
                    maxSubjectsPerRoom
                )
                    continue;

                // 分配學生
                ctx.students.forEach(function (student) {
                    if (student[columns.session] !== sessionNumber) return;
                    const studentKey =
                        student[columns.class] + student[columns.subject];
                    if (
                        studentKey === classSubjectKey &&
                        student[columns.room] === 0
                    ) {
                        student[columns.room] = roomNumber;
                        room.population++;
                        room.classSubject[classSubjectKey] =
                            (room.classSubject[classSubjectKey] || 0) + 1;
                    }
                });

                scheduledKeys.push(classSubjectKey);
            }
        }

        // 檢查未分配
        sessionStudents.forEach(function (student) {
            if (student[columns.room] === 0) {
                allScheduled = false;
            }
        });
    }

    if (!allScheduled) {
        ctx.warnings.push(
            "現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
    }

    // 檢查第 9 節
    const hasSession9 = ctx.students.some(function (s) {
        return s[columns.session] === 9;
    });
    if (hasSession9) {
        ctx.warnings.push(
            "部分考生被安排在第9節補考，請注意是否需要調整到中午應試！"
        );
    }

    ctx.stats.allScheduled = allScheduled;
    ctx.stats.hasSession9 = hasSession9;
    return ctx;
}

/**
 * 排序學生（依節次試場）
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_sortStudents(ctx) {
    const columns = ctx.columns;

    ctx.students.sort(function (a, b) {
        if (a[columns.session] !== b[columns.session])
            return a[columns.session] - b[columns.session];
        if (a[columns.room] !== b[columns.room])
            return a[columns.room] - b[columns.room];
        if (a[columns.department] !== b[columns.department])
            return a[columns.department].localeCompare(
                b[columns.department],
                "zh-TW"
            );
        if (a[columns.grade] !== b[columns.grade])
            return a[columns.grade].localeCompare(b[columns.grade], "zh-TW");
        if (a[columns.seatNumber] !== b[columns.seatNumber])
            return a[columns.seatNumber] - b[columns.seatNumber];
        return a[columns.subject].localeCompare(b[columns.subject], "zh-TW");
    });

    return ctx;
}

/**
 * 計算大、小袋編號
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_allocateBagIdentifiers(ctx) {
    const columns = ctx.columns;

    if (ctx.students.length === 0) return ctx;

    let smallBagCounter = 1;
    let bigBagCounter = 1;
    let lastSession = null;
    let lastRoom = null;
    let smallBagMapping = {};

    ctx.students.forEach(function (student) {
        const session = student[columns.session];
        const room = student[columns.room];
        const classSubjectKey =
            student[columns.class] + student[columns.subject];

        // 新試場 -> 新大袋
        if (session !== lastSession || room !== lastRoom) {
            if (lastSession !== null) {
                bigBagCounter++;
            }
            smallBagMapping = {};
            lastSession = session;
            lastRoom = room;
        }

        // 新班級科目組合 -> 新小袋
        if (!smallBagMapping[classSubjectKey]) {
            smallBagMapping[classSubjectKey] = smallBagCounter++;
        }

        student[columns.smallBagId] = smallBagMapping[classSubjectKey];
        student[columns.bigBagId] = bigBagCounter;
    });

    ctx.stats.totalSmallBags = smallBagCounter - 1;
    ctx.stats.totalBigBags = bigBagCounter;
    return ctx;
}

/**
 * 填入節次時間
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_populateSessionTimes(ctx) {
    const columns = ctx.columns;

    // 建立節次到時間的對映
    const timeLookup = {};
    ctx.sessionTimes.slice(1).forEach(function (row) {
        timeLookup[row[0]] = row[1];
    });

    ctx.students.forEach(function (student) {
        const session = student[columns.session];
        student[columns.time] = timeLookup[session] || "";
    });

    return ctx;
}

/**
 * 計算各種人數
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_updatePopulations(ctx) {
    const columns = ctx.columns;

    // 計算班級總人數
    const classPopulationMap = {};
    ctx.students.forEach(function (student) {
        const className = student[columns.class];
        classPopulationMap[className] =
            (classPopulationMap[className] || 0) + 1;
    });

    // 計算每個試場內的班級科目人數
    const roomClassSubjectMap = {};
    const roomPopulationMap = {};

    ctx.students.forEach(function (student) {
        const roomKey = student[columns.session] + "_" + student[columns.room];
        const classSubjectKey =
            roomKey + "_" + student[columns.class] + student[columns.subject];

        roomClassSubjectMap[classSubjectKey] =
            (roomClassSubjectMap[classSubjectKey] || 0) + 1;
        roomPopulationMap[roomKey] = (roomPopulationMap[roomKey] || 0) + 1;
    });

    // 更新學生資料
    ctx.students.forEach(function (student) {
        const roomKey = student[columns.session] + "_" + student[columns.room];
        const classSubjectKey =
            roomKey + "_" + student[columns.class] + student[columns.subject];

        student[columns.smallBagPopulation] =
            roomClassSubjectMap[classSubjectKey] || 0;
        student[columns.bigBagPopulation] = roomPopulationMap[roomKey] || 0;
        student[columns.classPopulation] =
            classPopulationMap[student[columns.class]] || 0;
    });

    ctx.stats.totalStudents = ctx.students.length;
    return ctx;
}

// ============================================================================
// Pipeline 組合執行器
// ============================================================================

/**
 * 執行完整排程 Pipeline
 *
 * 依序執行所有排程步驟：
 * 1. 安排共同科目節次
 * 2. 安排專業科目節次
 * 3. 安排試場
 * 4. 排序學生
 * 5. 計算大小袋編號
 * 6. 填入節次時間
 * 7. 計算各種人數
 *
 * @returns {Object} 執行結果，包含統計和警告
 */
function pipeline_executeFullScheduling() {
    const runtimeStart = new Date();

    // 階段 1: 建立候選名單
    buildFilteredCandidateList();

    // 階段 2: 載入資料
    let ctx = pipeline_loadData();

    // 階段 3: 連續處理（資料在記憶體中傳遞）
    ctx = pipeline_scheduleCommonSubjects(ctx);
    ctx = pipeline_scheduleSpecializedSubjects(ctx);
    ctx = pipeline_assignRooms(ctx);
    ctx = pipeline_sortStudents(ctx);
    ctx = pipeline_allocateBagIdentifiers(ctx);
    ctx = pipeline_populateSessionTimes(ctx);
    ctx = pipeline_updatePopulations(ctx);

    // 階段 4: 一次性寫回
    pipeline_saveData(ctx);

    // 階段 5: 產生報表
    createExamBulletinSheet();
    createProctorRecordSheet();
    composeSmallBagDataset();
    composeBigBagDataset();

    // 顯示警告
    ctx.warnings.forEach(function (warning) {
        SpreadsheetApp.getUi().alert(warning);
    });

    const elapsedSeconds = calculateElapsedSeconds(runtimeStart);
    ctx.stats.elapsedSeconds = elapsedSeconds;

    SpreadsheetApp.getUi().alert("已完成編排，共使用" + elapsedSeconds + "秒");

    return {
        stats: ctx.stats,
        warnings: ctx.warnings,
    };
}

/**
 * 執行手動調整後的 Pipeline
 *
 * 跳過節次和試場分配，只執行後續步驟：
 * 1. 排序學生
 * 2. 計算大小袋編號
 * 3. 填入節次時間
 * 4. 計算各種人數
 *
 * @returns {Object} 執行結果，包含統計和警告
 */
function pipeline_executePostAdjustment() {
    const runtimeStart = new Date();

    // 載入資料
    let ctx = pipeline_loadData();

    // 連續處理（資料在記憶體中傳遞）
    ctx = pipeline_sortStudents(ctx);
    ctx = pipeline_allocateBagIdentifiers(ctx);
    ctx = pipeline_populateSessionTimes(ctx);
    ctx = pipeline_updatePopulations(ctx);

    // 一次性寫回
    pipeline_saveData(ctx);

    // 產生報表
    createExamBulletinSheet();
    createProctorRecordSheet();
    composeSmallBagDataset();
    composeBigBagDataset();

    const elapsedSeconds = calculateElapsedSeconds(runtimeStart);
    ctx.stats.elapsedSeconds = elapsedSeconds;

    SpreadsheetApp.getUi().alert("已完成編排，共使用" + elapsedSeconds + "秒");

    return {
        stats: ctx.stats,
        warnings: ctx.warnings,
    };
}

/**
 * 只執行節次分配 Pipeline（不含試場分配）
 *
 * 適用於只需要分配節次的情境。
 *
 * @returns {Object} 執行結果
 */
function pipeline_executeSessionScheduling() {
    // 載入資料
    let ctx = pipeline_loadData();

    // 只執行節次分配
    ctx = pipeline_scheduleCommonSubjects(ctx);
    ctx = pipeline_scheduleSpecializedSubjects(ctx);

    // 寫回
    pipeline_saveData(ctx);

    // 顯示警告
    ctx.warnings.forEach(function (warning) {
        SpreadsheetApp.getUi().alert(warning);
    });

    return {
        stats: ctx.stats,
        warnings: ctx.warnings,
    };
}

/**
 * 只執行試場分配 Pipeline
 *
 * 適用於節次已分配完成，只需要分配試場的情境。
 *
 * @returns {Object} 執行結果
 */
function pipeline_executeRoomAssignment() {
    // 載入資料
    let ctx = pipeline_loadData();

    // 只執行試場分配
    ctx = pipeline_assignRooms(ctx);

    // 寫回
    pipeline_saveData(ctx);

    // 顯示警告
    ctx.warnings.forEach(function (warning) {
        SpreadsheetApp.getUi().alert(warning);
    });

    return {
        stats: ctx.stats,
        warnings: ctx.warnings,
    };
}
