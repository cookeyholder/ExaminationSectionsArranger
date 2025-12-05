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
    const columns = ctx.columns;
    const sessionRules = ctx.sessionRules;

    for (let i = 0; i < ctx.students.length; i++) {
        const student = ctx.students[i];
        const subjectName = student[columns.subject];
        const preferredSession = sessionRules[subjectName];
        if (preferredSession != null) {
            student[columns.session] = preferredSession;
        }
    }

    return ctx;
}

/**
 * 安排專業科目節次（優化版：使用預建索引）
 *
 * 優化重點：
 * 1. 預先建立「科別+年級+科目 → 學生索引陣列」的對映
 * 2. 分配時直接透過索引存取，避免遍歷整個學生陣列
 * 3. 時間複雜度從 O(n × m) 降為 O(n + m)
 *
 * @param {Object} ctx - Pipeline 上下文
 * @returns {Object} 更新後的上下文
 */
function pipeline_scheduleSpecializedSubjects(ctx) {
    const sessionCapacity = 0.9 * ctx.params.sessionCapacity;
    const maxSessionCount = ctx.params.maxSessionCount;
    const columns = ctx.columns;
    const students = ctx.students;

    // ===== 步驟 1: 預建索引（單次遍歷）=====
    // studentIndicesByGroup: "科別+年級+科目" -> [學生在陣列中的索引]
    const studentIndicesByGroup = {};
    const deptGradeSubjectCounts = {};

    for (let i = 0; i < students.length; i++) {
        const student = students[i];
        // 只處理尚未分配節次的學生
        if (student[columns.session] !== 0 && student[columns.session] !== "") {
            continue;
        }

        const key =
            student[columns.department] +
            student[columns.grade] +
            "_" +
            student[columns.subject];

        // 建立索引對映
        if (!studentIndicesByGroup[key]) {
            studentIndicesByGroup[key] = [];
        }
        studentIndicesByGroup[key].push(i);

        // 同時統計人數
        deptGradeSubjectCounts[key] = (deptGradeSubjectCounts[key] || 0) + 1;
    }

    // 依人數排序（大群組優先）
    const sortedCounts = Object.entries(deptGradeSubjectCounts).sort(
        compareCountDescending
    );

    // ===== 步驟 2: 節次統計初始化（計入已被共同科目佔用的人數）=====
    const sessionStats = {};
    for (let i = 1; i <= maxSessionCount; i++) {
        sessionStats[i] = { population: 0, deptGrade: {} };
    }

    // 統計已分配節次的學生（共同科目），更新 sessionStats
    for (let i = 0; i < students.length; i++) {
        const student = students[i];
        const sessionNum = student[columns.session];
        if (sessionNum > 0 && sessionNum <= maxSessionCount) {
            sessionStats[sessionNum].population++;

            // 同時記錄該科別年級已在此節次有科目
            const deptGradeKey =
                student[columns.department] + student[columns.grade];
            sessionStats[sessionNum].deptGrade[deptGradeKey] = true;
        }
    }

    // 記錄已分配的群組
    const assignedGroups = {};

    // ===== 步驟 3: 分配學生到節次（使用索引直接存取）=====
    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        const session = sessionStats[sessionNumber];

        for (let i = 0; i < sortedCounts.length; i++) {
            const [deptGradeSubjectKey, studentCount] = sortedCounts[i];

            // 已分配過的群組跳過
            if (assignedGroups[deptGradeSubjectKey]) {
                continue;
            }

            const deptGradeKey = deptGradeSubjectKey.substring(
                0,
                deptGradeSubjectKey.indexOf("_")
            );

            // 檢查互斥規則：同科別年級不能在同節有不同科目
            if (session.deptGrade[deptGradeKey]) {
                continue;
            }

            // 檢查容量
            if (studentCount + session.population > sessionCapacity) {
                continue;
            }

            // ===== 關鍵優化：透過索引直接存取學生 =====
            const indices = studentIndicesByGroup[deptGradeSubjectKey];
            for (let j = 0; j < indices.length; j++) {
                students[indices[j]][columns.session] = sessionNumber;
            }

            // 更新統計
            session.population += studentCount;
            session.deptGrade[deptGradeKey] =
                (session.deptGrade[deptGradeKey] || 0) + 1;

            // 標記已分配
            assignedGroups[deptGradeSubjectKey] = true;
        }
    }

    // ===== 步驟 4: 檢查未分配 =====
    let unscheduledCount = 0;
    for (let i = 0; i < students.length; i++) {
        const session = students[i][columns.session];
        if (session === 0 || session === "") {
            unscheduledCount++;
        }
    }

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
 * 安排試場（優化版：使用預建索引 + 最佳適配策略）
 *
 * 優化重點：
 * 1. 先依節次分組，避免每次迴圈都過濾整個學生陣列
 * 2. 預建「班級+科目 → 學生索引陣列」的對映
 * 3. 分配時直接透過索引存取，避免遍歷
 * 4. 採用「最佳適配」策略：為每個群組找最適合的試場
 * 5. 時間複雜度從 O(n × s × r) 降為 O(n + s × r)
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
    const students = ctx.students;

    // ===== 步驟 1: 重置所有試場並建立節次索引 =====
    // studentIndicesBySession: sessionNumber -> [學生索引陣列]
    const studentIndicesBySession = {};

    for (let i = 0; i < students.length; i++) {
        students[i][columns.room] = 0;

        const sessionNum = Number(students[i][columns.session]);
        if (sessionNum > 0) {
            if (!studentIndicesBySession[sessionNum]) {
                studentIndicesBySession[sessionNum] = [];
            }
            studentIndicesBySession[sessionNum].push(i);
        }
    }

    let allScheduled = true;

    // ===== 步驟 2: 逐節處理 =====
    for (
        let sessionNumber = 1;
        sessionNumber <= maxSessionCount;
        sessionNumber++
    ) {
        const sessionIndices = studentIndicesBySession[sessionNumber];
        if (!sessionIndices || sessionIndices.length === 0) continue;

        // ===== 步驟 2a: 建立該節次的「班級+科目 → 索引」對映 =====
        const indicesByClassSubject = {};
        const classSubjectCounts = {};

        for (let i = 0; i < sessionIndices.length; i++) {
            const studentIdx = sessionIndices[i];
            const student = students[studentIdx];
            const key = student[columns.class] + student[columns.subject];

            if (!indicesByClassSubject[key]) {
                indicesByClassSubject[key] = [];
            }
            indicesByClassSubject[key].push(studentIdx);
            classSubjectCounts[key] = (classSubjectCounts[key] || 0) + 1;
        }

        // 依人數排序（大群組優先）
        const sortedCounts = Object.entries(classSubjectCounts).sort(
            compareCountDescending
        );

        // ===== 步驟 2b: 試場統計初始化 =====
        const roomStats = {};
        for (let r = 1; r <= maxRoomCount; r++) {
            // subjects: 該試場中的科目集合（用於正確計算科目數）
            roomStats[r] = { population: 0, subjects: {} };
        }

        // ===== 步驟 2c: 採用「最佳適配」策略分配到試場 =====
        // 對每個群組，找出最適合的試場（剩餘空間最小但足夠容納的試場）
        for (let i = 0; i < sortedCounts.length; i++) {
            const [classSubjectKey, count] = sortedCounts[i];

            // 從 classSubjectKey 中提取科目名稱
            // classSubjectKey 格式為「班級 + 科目」，需要找到科目部分
            // 先從 indicesByClassSubject 中取一個學生來獲取科目名稱
            const sampleStudentIdx = indicesByClassSubject[classSubjectKey][0];
            const subjectName = students[sampleStudentIdx][columns.subject];

            // 找出最適合的試場
            let bestRoom = -1;
            let bestRemainingSpace = Infinity;

            for (let roomNumber = 1; roomNumber <= maxRoomCount; roomNumber++) {
                const room = roomStats[roomNumber];

                // 檢查人數限制
                if (count + room.population > maxStudentsPerRoom) continue;

                // 檢查科目數限制（只有新科目才需要檢查）
                const currentSubjectCount = Object.keys(room.subjects).length;
                const isNewSubject = !room.subjects[subjectName];
                if (
                    isNewSubject &&
                    currentSubjectCount + 1 > maxSubjectsPerRoom
                )
                    continue;

                // 計算剩餘空間，選擇剩餘空間最小的（最佳適配）
                const remainingSpace =
                    maxStudentsPerRoom - room.population - count;
                if (remainingSpace < bestRemainingSpace) {
                    bestRemainingSpace = remainingSpace;
                    bestRoom = roomNumber;
                }
            }

            // 如果找到適合的試場，進行分配
            if (bestRoom > 0) {
                const indices = indicesByClassSubject[classSubjectKey];
                for (let j = 0; j < indices.length; j++) {
                    students[indices[j]][columns.room] = bestRoom;
                }

                // 更新統計
                roomStats[bestRoom].population += count;
                roomStats[bestRoom].subjects[subjectName] = true;
            }
        }

        // ===== 步驟 2d: 檢查本節未分配 =====
        for (let i = 0; i < sessionIndices.length; i++) {
            if (students[sessionIndices[i]][columns.room] === 0) {
                allScheduled = false;
                break;
            }
        }
    }

    if (!allScheduled) {
        ctx.warnings.push(
            "現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
    }

    // 檢查第 9 節
    const hasSession9 =
        studentIndicesBySession[9] && studentIndicesBySession[9].length > 0;

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

    // 建立節次到時間的對映（使用字串鍵確保類型一致）
    const timeLookup = {};
    ctx.sessionTimes.slice(1).forEach(function (row) {
        timeLookup[String(row[0])] = row[1];
    });

    ctx.students.forEach(function (student) {
        const session = String(student[columns.session]);
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

    // 階段 5: 產生報表（直接使用記憶體中的資料，避免重複讀取和排序）
    pipeline_generateAllReports(ctx);

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

    // 產生報表（直接使用記憶體中的資料，避免重複讀取和排序）
    pipeline_generateAllReports(ctx);

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

// ============================================================================
// 優化版報表函式（使用記憶體中的資料）
// ============================================================================

/**
 * 按班級座號排序學生（記憶體內排序，不寫回工作表）
 *
 * 排序順序：科別 → 年級 → 座號 → 節次 → 科目
 *
 * @param {Array} students - 學生陣列
 * @param {Object} columns - 欄位索引
 * @returns {Array} 排序後的學生陣列副本
 */
function pipeline_sortByClassSeat(students, columns) {
    return students.slice().sort(function (a, b) {
        if (a[columns.department] !== b[columns.department])
            return a[columns.department].localeCompare(
                b[columns.department],
                "zh-TW"
            );
        if (a[columns.grade] !== b[columns.grade])
            return a[columns.grade].localeCompare(b[columns.grade], "zh-TW");
        if (a[columns.seatNumber] !== b[columns.seatNumber])
            return a[columns.seatNumber] - b[columns.seatNumber];
        if (a[columns.session] !== b[columns.session])
            return a[columns.session] - b[columns.session];
        return a[columns.subject].localeCompare(b[columns.subject], "zh-TW");
    });
}

/**
 * 產生公告版補考場次（使用記憶體中的資料）
 *
 * @param {Object} ctx - Pipeline 上下文
 */
function pipeline_createExamBulletinSheet(ctx) {
    const columns = ctx.columns;

    // 按班級座號排序（建立副本，不影響原資料）
    const sortedStudents = pipeline_sortByClassSeat(ctx.students, columns);

    BULLETIN_OUTPUT_SHEET.clear();
    if (BULLETIN_OUTPUT_SHEET.getMaxRows() > 5) {
        BULLETIN_OUTPUT_SHEET.deleteRows(
            2,
            BULLETIN_OUTPUT_SHEET.getMaxRows() - 5
        );
    }

    const bulletinRows = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
    sortedStudents.forEach(function (student) {
        const studentName = student[columns.name];
        let maskedName = "";
        if (studentName.length === 2) {
            maskedName = studentName.toString().slice(0, 1) + "〇";
        } else {
            const middleMaskLength = studentName.length - 2;
            maskedName =
                studentName.toString().slice(0, 1) +
                "〇".repeat(middleMaskLength) +
                studentName.toString().slice(-1);
        }

        bulletinRows.push([
            student[columns.class],
            student[columns.studentId],
            maskedName,
            student[columns.subject],
            student[columns.session],
            student[columns.room],
        ]);
    });

    writeRangeValuesSafely(
        BULLETIN_OUTPUT_SHEET.getRange(
            2,
            1,
            bulletinRows.length,
            bulletinRows[0].length
        ),
        bulletinRows
    );

    // 格式化
    const schoolYearValue =
        ctx.params.schoolYear || PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue =
        ctx.params.semester || PARAMETERS_SHEET.getRange("B3").getValue();

    BULLETIN_OUTPUT_SHEET.getRange("A1:F1").mergeAcross();
    BULLETIN_OUTPUT_SHEET.getRange("A1").setValue(
        "高雄高工" +
            schoolYearValue +
            "學年度第" +
            semesterValue +
            "學期補考名單"
    );
    BULLETIN_OUTPUT_SHEET.getRange("A1").setFontSize(20);
    BULLETIN_OUTPUT_SHEET.getRange(
        1,
        1,
        BULLETIN_OUTPUT_SHEET.getMaxRows(),
        BULLETIN_OUTPUT_SHEET.getMaxColumns()
    ).setHorizontalAlignment("center");
    BULLETIN_OUTPUT_SHEET.setFrozenRows(2);
    BULLETIN_OUTPUT_SHEET.getRange("A2:F").createFilter();
    BULLETIN_OUTPUT_SHEET.getRange("A2:F").setBorder(
        true,
        true,
        true,
        true,
        true,
        true,
        "#000000",
        SpreadsheetApp.BorderStyle.SOLID
    );
}

/**
 * 產生試場紀錄表（使用記憶體中的資料）
 *
 * @param {Object} ctx - Pipeline 上下文
 */
function pipeline_createProctorRecordSheet(ctx) {
    const columns = ctx.columns;
    const schoolYearValue =
        ctx.params.schoolYear || PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue =
        ctx.params.semester || PARAMETERS_SHEET.getRange("B3").getValue();

    // 資料已經按節次試場排序
    const students = ctx.students;

    RECORD_OUTPUT_SHEET.clear();
    if (RECORD_OUTPUT_SHEET.getMaxRows() > 5) {
        RECORD_OUTPUT_SHEET.deleteRows(2, RECORD_OUTPUT_SHEET.getMaxRows() - 5);
    }

    const recordRows = [
        [
            "A表：" +
                schoolYearValue +
                "學年度第" +
                semesterValue +
                "學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
            "",
        ],
        [
            "節次",
            "試場",
            "時間",
            "班級",
            "學號",
            "姓名",
            "科目名稱",
            "班級人數",
            "考生到考簽名",
            "違規記錄(打V)",
            "",
            "其他違規\n請簡述",
        ],
        ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""],
    ];

    students.forEach(function (student) {
        recordRows.push([
            student[columns.session],
            student[columns.room],
            student[columns.time],
            student[columns.class],
            student[columns.studentId],
            student[columns.name],
            student[columns.subject],
            // A 表「班級人數」欄位應為「同節次、同試場、同班級、同科目」的人數
            // 這與小袋的定義相同，所以使用 smallBagPopulation
            student[columns.smallBagPopulation],
            "",
            "",
            "",
            "",
        ]);
    });

    writeRangeValuesSafely(
        RECORD_OUTPUT_SHEET.getRange(
            1,
            1,
            recordRows.length,
            recordRows[0].length
        ),
        recordRows
    );

    // 格式化
    RECORD_OUTPUT_SHEET.getRange("A1:L1")
        .mergeAcross()
        .setVerticalAlignment("bottom")
        .setFontSize(14)
        .setFontWeight("bold");
    RECORD_OUTPUT_SHEET.getRange("J2:K2").mergeAcross();
    RECORD_OUTPUT_SHEET.getRange("A2:A3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("B2:B3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("C2:C3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("D2:D3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("E2:E3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("F2:F3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("G2:G3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("H2:H3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("I2:I3")
        .mergeVertically()
        .setVerticalAlignment("middle");
    RECORD_OUTPUT_SHEET.getRange("L2:L3")
        .mergeVertically()
        .setVerticalAlignment("middle");

    RECORD_OUTPUT_SHEET.getRange(
        2,
        1,
        recordRows.length + 2,
        recordRows[0].length
    )
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setBorder(
            true,
            true,
            true,
            true,
            true,
            true,
            "#000000",
            SpreadsheetApp.BorderStyle.SOLID
        );
}

/**
 * 產生小袋封面套印用資料（使用記憶體中的資料）
 *
 * @param {Object} ctx - Pipeline 上下文
 */
function pipeline_composeSmallBagDataset(ctx) {
    const columns = ctx.columns;
    const schoolYearValue =
        ctx.params.schoolYear || PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue =
        ctx.params.semester || PARAMETERS_SHEET.getRange("B3").getValue();

    SMALL_BAG_DATA_SHEET.clear();
    if (SMALL_BAG_DATA_SHEET.getMaxRows() > 5) {
        SMALL_BAG_DATA_SHEET.deleteRows(
            2,
            SMALL_BAG_DATA_SHEET.getMaxRows() - 5
        );
    }

    const smallBagRows = [
        [
            "學年度",
            "學期",
            "小袋序號",
            "節次",
            "時間",
            "試場",
            "班級",
            "科目名稱",
            "任課老師",
            "小袋人數",
            "電腦",
            "人工",
        ],
    ];
    const processedSmallBags = [];

    ctx.students.forEach(function (student) {
        const smallBagId = student[columns.smallBagId];
        if (!processedSmallBags.includes(smallBagId)) {
            smallBagRows.push([
                schoolYearValue,
                semesterValue,
                smallBagId,
                student[columns.session],
                student[columns.time],
                student[columns.room],
                student[columns.class],
                student[columns.subject],
                student[columns.teacher],
                student[columns.smallBagPopulation],
                student[columns.computer],
                student[columns.manual],
            ]);
            processedSmallBags.push(smallBagId);
        }
    });

    writeRangeValuesSafely(
        SMALL_BAG_DATA_SHEET.getRange(
            1,
            1,
            smallBagRows.length,
            smallBagRows[0].length
        ),
        smallBagRows
    );
}

/**
 * 產生大袋封面套印用資料（使用記憶體中的資料）
 *
 * @param {Object} ctx - Pipeline 上下文
 */
function pipeline_composeBigBagDataset(ctx) {
    const columns = ctx.columns;
    const schoolYearValue =
        ctx.params.schoolYear || PARAMETERS_SHEET.getRange("B2").getValue();
    const semesterValue =
        ctx.params.semester || PARAMETERS_SHEET.getRange("B3").getValue();
    const makeUpDateValue = PARAMETERS_SHEET.getRange("B13").getValue();
    const invigilatorAssignment = transposeMatrix(
        INVIGILATOR_ASSIGNMENT_SHEET.getDataRange().getValues()
    );

    BIG_BAG_DATA_SHEET.clear();
    if (BIG_BAG_DATA_SHEET.getMaxRows() > 5) {
        BIG_BAG_DATA_SHEET.deleteRows(2, BIG_BAG_DATA_SHEET.getMaxRows() - 5);
    }

    const bigBagRows = [
        [
            "學年度",
            "學期",
            "大袋序號",
            "節次",
            "試場",
            "補考日期",
            "時間",
            "試卷袋序號",
            "監考教師",
            "各試場人數",
        ],
    ];
    const processedBigBags = [];

    // 計算每個大袋包含的小袋範圍
    const smallBagRangeByBigBag = {};
    ctx.students.forEach(function (student) {
        const bigBagKey = "大袋" + student[columns.bigBagId];
        if (!Object.keys(smallBagRangeByBigBag).includes(bigBagKey)) {
            smallBagRangeByBigBag[bigBagKey] = [student[columns.smallBagId]];
        } else {
            smallBagRangeByBigBag[bigBagKey].push(student[columns.smallBagId]);
        }
    });

    ctx.students.forEach(function (student) {
        const bigBagId = student[columns.bigBagId];
        if (!processedBigBags.includes(bigBagId)) {
            const sessionNum = parseInt(student[columns.session]);
            const roomNum = parseInt(student[columns.room]);

            if (isNaN(sessionNum) || isNaN(roomNum)) {
                Logger.log(
                    `無效的節次/試場索引: 節次=${sessionNum}, 試場=${roomNum}`
                );
                return;
            }

            const bagRange = smallBagRangeByBigBag["大袋" + bigBagId];

            var invigilatorName = "";
            if (invigilatorAssignment && invigilatorAssignment[sessionNum]) {
                invigilatorName =
                    invigilatorAssignment[sessionNum][roomNum] || "";
            }

            bigBagRows.push([
                schoolYearValue,
                semesterValue,
                bigBagId,
                student[columns.session],
                student[columns.room],
                makeUpDateValue,
                student[columns.time],
                Math.min(...bagRange) + "-" + Math.max(...bagRange),
                invigilatorName,
                student[columns.bigBagPopulation],
            ]);

            processedBigBags.push(bigBagId);
        }
    });

    writeRangeValuesSafely(
        BIG_BAG_DATA_SHEET.getRange(
            1,
            1,
            bigBagRows.length,
            bigBagRows[0].length
        ),
        bigBagRows
    );
}

/**
 * 產生所有報表（使用記憶體中的資料）
 *
 * @param {Object} ctx - Pipeline 上下文
 */
function pipeline_generateAllReports(ctx) {
    pipeline_createExamBulletinSheet(ctx);
    pipeline_createProctorRecordSheet(ctx);
    pipeline_composeSmallBagDataset(ctx);
    pipeline_composeBigBagDataset(ctx);
}
