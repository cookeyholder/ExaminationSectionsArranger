// 建立工作列選單
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
  const headers = ["科別", "年級", "班級代碼", "班級","座號",	"學號",	"姓名","科目名稱","節次", "試場", "小袋序號", "小袋人數", "大袋序號", "大袋人數", "班級人數", "時間", "電腦", "人工", "任課老師"]
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
  let runtimeCountStart = new Date();

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

function runtimeCountStop(start) {

  const stop = new Date();
  const newRuntime = Number(stop) - Number(start);
  return Math.ceil(newRuntime/1000);

}


function countTimeConsume(runner){
  const startTime = new Date();
  runner()
  const endTime = new Date()
  const runtime = Math.ceil(Number(endTime) - Number(startTime))/1000;
  return runtime;
}


function setRangeValues(range, data){
  if (range.getLastColumn() == data[0].length){
    range.setValues(data);
  }
  else {
    SpreadsheetApp.getUi().alert("欲寫入範圍欄數不足！");
  }
}


function getClassCode(cls){
  // 班級代碼查詢
  // 輸入班級中文名稱(4字)，輸出6碼數字代碼
  // 如輸入「機械二丁」，輸出「301204」。

  const departmentCodes = {
    "機械": "301",
    "汽車": "303",
    "資訊": "305",
    "電子": "306",
    "電機": "308",
    "冷凍": "309",
    "建築": "311",
    "化工": "315",
    "圖傳": "373",
    "電圖": "374"
  }

  const classAndGradeCode = {
    "甲": "01",
    "乙": "02",
    "丙": "03",
    "丁": "04",
    "一": "1",
    "二": "2",
    "三": "3"
  };

  return departmentCodes[cls.slice(0,2)] + classAndGradeCode[cls.slice(2,3)] + classAndGradeCode[cls.slice(-1)]
}


function getGrade(cls){
  const grade = {
    "一": "1",
    "二": "2",
    "三": "3"
  };

  return grade[cls.slice(2,3)]
}


function getDepartmentName(cls){
  const departments = {
    "機械": "機械科",
    "汽車": "汽車科",
    "資訊": "資訊科",
    "電子": "電子科",
    "電機": "電機科",
    "冷凍": "冷凍空調科",
    "建築": "建築科",
    "化工": "化工科",
    "圖傳": "圖文傳播科",
    "電圖": "電腦機械製圖科"
  }

  return departments[cls.slice(0,2)];
}


function checkShowedBoxes(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const dataRange = ss.getSheetByName("教學組排入考程的科目").getRange("A2:A");
  const dataValues = dataRange.getValues();
  const numRows = dataRange.getNumRows();
  const numCols = dataRange.getNumColumns();

  for(let i = 0; i < numRows; i++) {
    if(!ss.isRowHiddenByFilter(i+1)){
      for(let j = 0; j < numCols; j++){
        dataValues[i][j] = true;
      }
    }
  }

  setRangeValues(dataRange, dataValues);
}


function cancelCheckboxes(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const dataRange = ss.getSheetByName("教學組排入考程的科目").getRange("A1:A");
  const dataValues = dataRange.getValues();
  const numRows = dataRange.getNumRows();
  const numCols = dataRange.getNumColumns();

  for(let i = 0; i < numRows; i++) {
    if(!ss.isRowHiddenByFilter(i+1)){
      for(let j = 0; j < numCols; j++){
        dataValues[i][j] = false;
      }
    }
  }

  setRangeValues(dataRange, dataValues);
}


function descendingPopulation(a, b) {
    if (a[1] === b[1]) {
        return 0;
    }
    else {
        return (a[1] < b[1]) ? 1 : -1;
    }
}


function getDepartmentGradeStatisticsOfArray(data){
  let departmentColumn = 0;
  let gradeColumn = 1;
  let statistics = {};
  for (const row of data){
    let key = row[departmentColumn] + row[gradeColumn];

    if (key in statistics){
      statistics[key] += 1;
    } else {
      statistics[key] = 1;
    }
  }

  return statistics;
}


function getDepartmentGradeSubjectStatisticsOfArray(data){
  const departmentColumn = 0;
  const gradeColumn = 1;
  const subjectNameColumn = 7;
  
  let statistics = {};
  for (const row of data){
    let key = row[departmentColumn] + row[gradeColumn] + "_" + row[subjectNameColumn];

    if (key in statistics){
      statistics[key] += 1;
    } else {
      statistics[key] = 1;
    }
  }

  return statistics;
}


function getDepartmentGradeStatistics(){
  // 統計各科別年級、各班級的應考人數
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();

  return getDepartmentGradeStatisticsOfArray(data);
}


function getDepartmentGradeSubjectStatistics(){
  // 統計各科別年級、各班級、科目的應考人數
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();

  return getDepartmentGradeSubjectStatisticsOfArray(data);
}


function createClassroom(){
  return {
    students: [],
    get population(){return this.students.length;},
    get classSubjectStatistics(){
      let statistics = {};
      this.students.forEach(
        function(row){
          let key = row[3] + "_" + row[7];  // 班級 + _ + 科目
          if(Object.keys(statistics).includes(key)){
            statistics[key] += 1;
          } else {
            statistics[key] = 1;
          }
        }
      );
      return statistics;
    }
  }
}

function createSession(){
  // session 物件工廠，用來產生下面的 getSessionStatistics 函數中，需要建立 9 個 session 物件
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_CLASSROOM_NUMBER = parametersSheet.getRange("B6").getValue();
  const session = {
    classrooms: [],
    students:[], 
    get population(){return this.students.length;},

    get departmentGradeStatistics(){
      let statistics = {};
      this.students.forEach(
        function(row){
          let key = row[0]+row[1];  // 科別 + 年級
          if(Object.keys(statistics).includes(key)){
            statistics[key] += 1;
          } else {
            statistics[key] = 1;
          }
        }
      );
      return statistics;
    },

    get departmentClassSubjectStatistics(){
      let statistics = {};
      this.students.forEach(
        function(row){
          let key = row[3] + row[7];  // 班級 + 科目
          if(Object.keys(statistics).includes(key)){
            statistics[key] += 1;
          } else {
            statistics[key] = 1;
          }
        }
      );
      return statistics;
    }
  };

  for(let j=0; j < MAX_CLASSROOM_NUMBER + 1; j++){
    session.classrooms.push(createClassroom())
  }
  return session;
}


function getSessionStatistics(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const sessionColumn = headers.indexOf("節次");
  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();

  const sessions = [];
  for(let i=0; i < MAX_SESSION_NUMBER + 2; i++){
    sessions.push(createSession())
  }

  for(const row of data){
    sessions[row[sessionColumn]].students.push(row);
  }

  return sessions;
}
