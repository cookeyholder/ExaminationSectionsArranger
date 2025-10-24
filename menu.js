function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("補考節次試場編排小幫手")
        .addItem("註冊組補考名單課程代碼補完", "updateUnfilteredSubjectCodes")
        .addItem("開課資料課程代碼補完", "updateOpenCourseCodes")
      .addSeparator()
        .addItem("步驟 1. 產出公告用補考名單、試場記錄表", "runFullSchedulingPipeline")
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "mergeSmallBagPdfFiles")
        .addItem("產生合併大袋封面用資料(人工輸入監考教師)", "composeBigBagDataset")
        .addItem("步驟 3. 合併列印大袋封面", "mergeBigBagPdfFiles")
      .addSeparator()
        .addItem("手動調整試場之後，繼續進行到試場紀錄表", "resumePipelineAfterManualAdjustments")
      .addSeparator()
        .addItem("依「科目」排序補考名單", "sortFilteredStudentsBySubject")
        .addItem("依「班級座號」排序補考名單", "sortFilteredStudentsByClassSeat")
        .addItem("依「節次試場」排序補考名單", "sortFilteredStudentsBySessionRoom")
      .addSeparator()
        .addItem("步驟 1-1. 清空", "resetFilteredSheets")
        .addItem("步驟 1-2. 開始篩選", "buildFilteredCandidateList")
        .addItem("步驟 1-3. 安排共同科節次", "scheduleCommonSubjectSessions")
        .addItem("步驟 1-4. 安排專業科節次", "scheduleSpecializedSubjectSessions")
        .addItem("步驟 1-5. 安排試場", "assignExamRooms")
        .addItem("步驟 1-6. 計算大、小袋編號", "allocateBagIdentifiers")
        .addItem("步驟 1-7. 填入試場時間", "populateSessionTimes")
        .addItem("步驟 1-8. 計算試場人數", "updateBagAndClassPopulations")
        .addItem("步驟 1-9. 產生「公告版補考場次」", "createExamBulletinSheet")
        .addItem("步驟 1-10. 產生「試場記錄表」", "createProctorRecordSheet")
        .addItem("步驟 1-11. 產生「小袋封面套印用資料」", "composeSmallBagDataset")
        .addItem("步驟 1-12. 產生「大袋封面套印用資料」", "composeBigBagDataset")
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "mergeSmallBagPdfFiles")
        .addItem("步驟 3. 合併列印大袋封面", "mergeBigBagPdfFiles")
    .addToUi();
}


function resetFilteredSheets(){
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
  const headers = ["科別", "年級", "班級代碼", "班級","座號",	"學號",	"姓名","科目名稱","節次", "試場", "小袋序號", "小袋人數", "大袋序號", "大袋人數", "班級人數", "時間", "電腦", "人工", "任課老師"];
  FILTERED_RESULT_SHEET.clear();
  FILTERED_RESULT_SHEET.appendRow(headers);

  // 移除已有篩選器，重新設置新的篩選器
  if(FILTERED_RESULT_SHEET.getDataRange().getFilter()){
    FILTERED_RESULT_SHEET.getDataRange().getFilter().remove();
  }
  FILTERED_RESULT_SHEET.getDataRange().createFilter();
}


// 一鍵產出公告用補考名單、試場記錄表
function runFullSchedulingPipeline(){
  // Start counting execution time
  const runtimeCountStart = new Date();

  buildFilteredCandidateList();  // 篩選出列入考程的科目
  scheduleCommonSubjectSessions();  // 安排物理、國、英、數、資訊科技、史地的節次
  scheduleSpecializedSubjectSessions();  // 安排專業科目的節次
  assignExamRooms();  // 安排試場的班級科目
  sortFilteredStudentsBySessionRoom();
  allocateBagIdentifiers();
  populateSessionTimes();
  updateBagAndClassPopulations();
  createExamBulletinSheet();
  createProctorRecordSheet();
  composeSmallBagDataset();
  composeBigBagDataset();

  const newRuntime = calculateElapsedSeconds(runtimeCountStart);

  SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}


function resumePipelineAfterManualAdjustments(){
  const runtimeCountStart = new Date();

  sortFilteredStudentsBySessionRoom();
  allocateBagIdentifiers();
  populateSessionTimes();
  updateBagAndClassPopulations();
  createExamBulletinSheet();
  createProctorRecordSheet();
  composeSmallBagDataset();
  composeBigBagDataset();

  const newRuntime = calculateElapsedSeconds(runtimeCountStart);

  SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}
