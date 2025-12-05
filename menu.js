function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu("補考節次試場編排小幫手")
        .addItem("註冊組補考名單課程代碼補完", "updateUnfilteredSubjectCodes")
        .addItem("開課資料課程代碼補完", "updateOpenCourseCodes")
        .addSeparator()
        .addItem(
            "步驟 1. 產出公告用補考名單、試場記錄表",
            "pipeline_executeFullScheduling"
        )
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "mergeSmallBagPdfFiles")
        .addItem(
            "產生合併大袋封面用資料(人工輸入監考教師)",
            "composeBigBagDataset"
        )
        .addItem("步驟 3. 合併列印大袋封面", "mergeBigBagPdfFiles")
        .addSeparator()
        .addItem(
            "手動調整試場之後，繼續進行到試場紀錄表",
            "pipeline_executePostAdjustment"
        )
        .addSeparator()
        .addItem("依「科目」排序補考名單", "sortFilteredStudentsBySubject")
        .addItem(
            "依「班級座號」排序補考名單",
            "sortFilteredStudentsByClassSeat"
        )
        .addItem(
            "依「節次試場」排序補考名單",
            "sortFilteredStudentsBySessionRoom"
        )
        .addSeparator()
        .addSubMenu(
            SpreadsheetApp.getUi()
                .createMenu("Pipeline 模式（高效能）")
                .addItem("完整排程", "pipeline_executeFullScheduling")
                .addItem("手動調整後繼續", "pipeline_executePostAdjustment")
                .addItem("只分配節次", "pipeline_executeSessionScheduling")
                .addItem("只分配試場", "pipeline_executeRoomAssignment")
        )
        .addSubMenu(
            SpreadsheetApp.getUi()
                .createMenu("單步驟模式（除錯用）")
                .addItem("步驟 1-1. 清空", "resetFilteredSheets")
                .addItem("步驟 1-2. 開始篩選", "buildFilteredCandidateList")
                .addItem(
                    "步驟 1-3. 安排共同科節次",
                    "scheduleCommonSubjectSessions"
                )
                .addItem(
                    "步驟 1-4. 安排專業科節次",
                    "scheduleSpecializedSubjectSessions"
                )
                .addItem("步驟 1-5. 安排試場", "assignExamRooms")
                .addItem("步驟 1-6. 計算大、小袋編號", "allocateBagIdentifiers")
                .addItem("步驟 1-7. 填入試場時間", "populateSessionTimes")
                .addItem(
                    "步驟 1-8. 計算試場人數",
                    "updateBagAndClassPopulations"
                )
                .addItem(
                    "步驟 1-9. 產生「公告版補考場次」",
                    "createExamBulletinSheet"
                )
                .addItem(
                    "步驟 1-10. 產生「試場記錄表」",
                    "createProctorRecordSheet"
                )
                .addItem(
                    "步驟 1-11. 產生「小袋封面套印用資料」",
                    "composeSmallBagDataset"
                )
                .addItem(
                    "步驟 1-12. 產生「大袋封面套印用資料」",
                    "composeBigBagDataset"
                )
        )
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "mergeSmallBagPdfFiles")
        .addItem("步驟 3. 合併列印大袋封面", "mergeBigBagPdfFiles")
        .addToUi();
}

function resetFilteredSheets() {
    // 將工作表「排入考程的補考名單」初始化成只剩第一列的欄位標題
    // (1) 清除所有儲存格內容
    // (2) 刪除多餘的列到只剩5列
    // (3) 填入欄位標題

    // 清除所有值
    FILTERED_RESULT_SHEET.clear();
    SMALL_BAG_DATA_SHEET.clear();
    BIG_BAG_DATA_SHEET.clear();
    BULLETIN_OUTPUT_SHEET.clear();
    RECORD_OUTPUT_SHEET.clear();

    // 將課程代碼補完，包括：「註冊組匯出的補考名單」、「開課資料(查詢任課教師用)」
    updateUnfilteredSubjectCodes();
    updateOpenCourseCodes();

    // 清空資料並設置標題列
    const headers = [
        "科別",
        "年級",
        "班級代碼",
        "班級",
        "座號",
        "學號",
        "姓名",
        "科目名稱",
        "節次",
        "試場",
        "小袋序號",
        "小袋人數",
        "大袋序號",
        "大袋人數",
        "班級人數",
        "時間",
        "電腦",
        "人工",
        "任課老師",
    ];
    FILTERED_RESULT_SHEET.clear();
    FILTERED_RESULT_SHEET.appendRow(headers);

    // 移除已有篩選器，重新設置新的篩選器
    if (FILTERED_RESULT_SHEET.getDataRange().getFilter()) {
        FILTERED_RESULT_SHEET.getDataRange().getFilter().remove();
    }
    FILTERED_RESULT_SHEET.getDataRange().createFilter();
}

// 一鍵產出公告用補考名單、試場記錄表
// 採用「單次讀取-Pipeline處理-單次寫入」模式優化效能
function runFullSchedulingPipeline() {
    const runtimeCountStart = new Date();

    // ===== 階段 1: 建立候選名單 =====
    buildFilteredCandidateList();

    // ===== 階段 2: 一次性讀取所有資料 =====
    const rawData = loadAllData();
    const students = rawData.students.slice(1); // 去除標題列
    const headerRow = rawData.students[0];
    const params = parseParameters(rawData.parameters);
    const columns = buildColumnIndices(headerRow);
    const sessionRules = parseSessionRules(rawData.sessionRules);

    // ===== 階段 3: Pipeline 處理（純函式，零 I/O）=====
    scheduleCommonSubjectSessionsInternal(students, columns, sessionRules);

    const unscheduled = scheduleSpecializedSubjectSessionsInternal(
        students,
        columns,
        params
    );
    if (unscheduled > 0) {
        SpreadsheetApp.getUi().alert(
            "無法將所有人排入 " +
                params.maxSessionCount +
                " 節，請檢查是否有某科年級須補考過多科目！"
        );
    }

    const allScheduled = assignExamRoomsInternal(students, columns, params);
    if (!allScheduled) {
        SpreadsheetApp.getUi().alert(
            "現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！"
        );
    }

    // 檢查第 9 節
    const hasSession9 = students.some(function (s) {
        return s[columns.session] === 9;
    });
    if (hasSession9) {
        SpreadsheetApp.getUi().alert(
            "部分考生被安排在第9節補考，請注意是否需要調整到中午應試！"
        );
    }

    sortStudentsInternal(students, columns);
    allocateBagIdentifiersInternal(students, columns);
    populateSessionTimesInternal(students, columns, rawData.sessionTimes);
    updateBagAndClassPopulationsInternal(students, columns);

    // ===== 階段 4: 一次性寫回 =====
    saveAllData(students);
    FILTERED_RESULT_SHEET.getRange("I:J").setNumberFormat("#,##0");

    // ===== 階段 5: 產生報表（需要讀取已排序的資料）=====
    createExamBulletinSheet();
    createProctorRecordSheet();
    composeSmallBagDataset();
    composeBigBagDataset();

    const newRuntime = calculateElapsedSeconds(runtimeCountStart);
    SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}

function resumePipelineAfterManualAdjustments() {
    const runtimeCountStart = new Date();

    // ===== 一次性讀取 =====
    const rawData = loadAllData();
    const students = rawData.students.slice(1);
    const headerRow = rawData.students[0];
    const columns = buildColumnIndices(headerRow);

    // ===== Pipeline 處理 =====
    sortStudentsInternal(students, columns);
    allocateBagIdentifiersInternal(students, columns);
    populateSessionTimesInternal(students, columns, rawData.sessionTimes);
    updateBagAndClassPopulationsInternal(students, columns);

    // ===== 一次性寫回 =====
    saveAllData(students);

    // ===== 產生報表 =====
    createExamBulletinSheet();
    createProctorRecordSheet();
    composeSmallBagDataset();
    composeBigBagDataset();

    const newRuntime = calculateElapsedSeconds(runtimeCountStart);
    SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}
