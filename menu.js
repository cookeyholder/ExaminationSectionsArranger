function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("補考節次試場編排小幫手")
        .addItem("註冊組補考名單課程代碼補完", "unfilteredCodeComplete")
        .addItem("開課資料課程代碼補完", "openCodeComplete")
      .addSeparator()
        .addItem("步驟 1. 產出公告用補考名單、試場記錄表", "allInOne")
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "mergeToSmallBag")
        .addItem("產生合併大袋封面用資料(人工輸入監考教師)", "generateBigBagData")
        .addItem("步驟 3. 合併列印大袋封面", "mergeToBigBag")
      .addSeparator()
        .addItem("手動調整試場之後，繼續進行到試場紀錄表", "afterRearrangeClassroomByManual")
      .addSeparator()
        .addItem("依「科目」排序補考名單", "sortBySubject")
        .addItem("依「班級座號」排序補考名單", "sortByClassname")
        .addItem("依「節次試場」排序補考名單", "sortBySessionClassroom")
      .addSeparator()
        .addItem("步驟 1-1. 清空", "initialize")
        .addItem("步驟 1-2. 開始篩選", "getFilteredData")
        .addItem("步驟 1-3. 安排共同科節次", "arrangeCommonsSession")
        .addItem("步驟 1-4. 安排專業科節次", "arrangeProfessionsSession")
        .addItem("步驟 1-5. 安排試場", "arrangeClassroom")
        .addItem("步驟 1-6. 計算大、小袋編號", "bagNumbering")
        .addItem("步驟 1-7. 填入試場時間", "setSessionTime")
        .addItem("步驟 1-8. 計算試場人數", "calculateClassroomPopulation")
        .addItem("步驟 1-9. 產生「公告版補考場次」", "generateBulletin")
        .addItem("步驟 1-10. 產生「試場記錄表」", "generateRecordSheet")
        .addItem("步驟 1-11. 產生「小袋封面套印用資料」", "generateSmallBagData")
        .addItem("步驟 1-12. 產生「大袋封面套印用資料」", "generateBigBagData")
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "mergeToSmallBag")
        .addItem("步驟 3. 合併列印大袋封面", "mergeToBigBag")
    .addToUi();
}


function initialize(){
  // 將工作表「排入考程的補考名單」初始化成只剩第一列的欄位標題
  // (1) 清除所有儲存格內容
  // (2) 刪除多餘的列到只剩5列
  // (3) 填入欄位標題

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const smallBagSheet = ss.getSheetByName("小袋封面套印用資料");
  const bigBagSheet = ss.getSheetByName("大袋封面套印用資料");
  const bulletinSheet = ss.getSheetByName("公告版補考場次");
  const recordSheet = ss.getSheetByName("試場紀錄表(A表)");
  
  // 清除所有值
  filteredSheet.clear();
  smallBagSheet.clear();
  bigBagSheet.clear();
  bulletinSheet.clear();
  recordSheet.clear();

  // 將課程代碼補完，包括：「註冊組匯出的補考名單」、「開課資料(查詢任課教師用)」
  unfilteredCodeComplete();
  openCodeComplete();
  

  // 清空資料並設置標題列
  const headers = ["科別", "年級", "班級代碼", "班級","座號",	"學號",	"姓名","科目名稱","節次", "試場", "小袋序號", "小袋人數", "大袋序號", "大袋人數", "班級人數", "時間", "電腦", "人工", "任課老師"];
  filteredSheet.clear();
  filteredSheet.appendRow(headers);

  // 移除已有篩選器，重新設置新的篩選器
  if(filteredSheet.getDataRange().getFilter()){
    filteredSheet.getDataRange().getFilter().remove();
  }
  filteredSheet.getDataRange().createFilter();
}


// 一鍵產出公告用補考名單、試場記錄表
function allInOne(){
  // Start counting execution time
  const runtimeCountStart = new Date();

  getFilteredData();  // 篩選出列入考程的科目
  arrangeCommonsSession();  // 安排物理、國、英、數、資訊科技、史地的節次
  arrangeProfessionsSession();  // 安排專業科目的節次
  arrangeClassroom();  // 安排試場的班級科目
  sortBySessionClassroom();
  bagNumbering();
  setSessionTime();
  calculateClassroomPopulation();
  generateBulletin();
  generateRecordSheet();
  generateSmallBagData();
  generateBigBagData();

  // Stop counting execution time
  const newRuntime = runtimeCountStop(runtimeCountStart);

  SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}


function afterRearrangeClassroomByManual(){
  const runtimeCountStart = new Date();

  sortBySessionClassroom();
  bagNumbering();
  setSessionTime();
  calculateClassroomPopulation();
  generateBulletin();
  generateRecordSheet();
  generateSmallBagData();
  generateBigBagData();

  const newRuntime = runtimeCountStop(runtimeCountStart);

  SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}
