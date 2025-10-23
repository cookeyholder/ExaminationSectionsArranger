// 建立工作列選單
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("補考節次試場編排小幫手")
        .addItem("註冊組補考名單課程代碼補完", "unfilted_code_complete")
        .addItem("開課資料課程代碼補完", "open_code_complete")
      .addSeparator()
        .addItem("步驟 1. 產出公告用補考名單、試場記錄表", "all_in_one")
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "merge_to_small_bag")
        .addItem("產生合併大袋封面用資料(人工輸入監考教師)", "generate_big_bag_data")
        .addItem("步驟 3. 合併列印大袋封面", "merge_to_big_bag")
      .addSeparator()
        .addItem("手動調整試場之後，繼續進行到試場紀錄表", "afterRearrangeClassroomByManual")
      .addSeparator()
        .addItem("依「科目」排序補考名單", "sort_by_subject")
        .addItem("依「班級座號」排序補考名單", "sort_by_classname")
        .addItem("依「節次試場」排序補考名單", "sort_by_session_classroom")
      .addSeparator()
        .addItem("步驟 1-1. 清空", "initialize")
        .addItem("步驟 1-2. 開始篩選", "get_filtered_data")
        .addItem("步驟 1-3. 安排共同科節次", "arrange_commons_session")
        .addItem("步驟 1-4. 安排專業科節次", "arrangeProfessionsSession")
        .addItem("步驟 1-5. 安排試場", "arrangeClassroom")
        .addItem("步驟 1-6. 計算大、小袋編號", "bag_numbering")
        .addItem("步驟 1-7. 填入試場時間", "set_session_time")
        .addItem("步驟 1-8. 計算試場人數", "calculate_classroom_population")
        .addItem("步驟 1-9. 產生「公告版補考場次」", "generate_bulletin")
        .addItem("步驟 1-10. 產生「試場記錄表」", "generate_record_sheet")
        .addItem("步驟 1-11. 產生「小袋封面套印用資料」", "generate_small_bag_data")
        .addItem("步驟 1-12. 產生「大袋封面套印用資料」", "generate_big_bag_data")
        .addItem("步驟 2. 合併列印小袋封面(要很久哦)", "merge_to_small_bag")
        .addItem("步驟 3. 合併列印大袋封面", "merge_to_big_bag")
    .addToUi();
}


function initialize(){
  // 將工作表「排入考程的補考名單」初始化成只剩第一列的欄位標題
  // (1) 清除所有儲存格內容
  // (2) 刪除多餘的列到只剩5列
  // (3) 填入欄位標題

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const small_bag_sheet = ss.getSheetByName("小袋封面套印用資料");
  const big_bag_sheet = ss.getSheetByName("大袋封面套印用資料");
  const bulletin_sheet = ss.getSheetByName("公告版補考場次");
  const record_sheet = ss.getSheetByName("試場紀錄表(A表)");
  
  // 清除所有值
  filtered_sheet.clear();
  small_bag_sheet.clear();
  big_bag_sheet.clear();
  bulletin_sheet.clear();
  record_sheet.clear();

  // 將課程代碼補完，包括：「註冊組匯出的補考名單」、「開課資料(查詢任課教師用)」
  unfilted_code_complete();
  open_code_complete();
  

  // 清空資料並設置標題列
  const headers = ["科別", "年級", "班級代碼", "班級","座號",	"學號",	"姓名","科目名稱","節次", "試場", "小袋序號", "小袋人數", "大袋序號", "大袋人數", "班級人數", "時間", "電腦", "人工", "任課老師"]
  filtered_sheet.clear();
  filtered_sheet.appendRow(headers);

  // 移除已有篩選器，重新設置新的篩選器
  if(filtered_sheet.getDataRange().getFilter()){
    filtered_sheet.getDataRange().getFilter().remove();
  }
  filtered_sheet.getDataRange().createFilter();
}


// 一鍵產出公告用補考名單、試場記錄表
function all_in_one(){
  // Start counting execution time
  var runtime_count_start = new Date();

  get_filtered_data();  // 篩選出列入考程的科目
  arrange_commons_session();  // 安排物理、國、英、數、資訊科技、史地的節次
  arrangeProfessionsSession();  // 安排專業科目的節次
  arrangeClassroom();  // 安排試場的班級科目
  sort_by_session_classroom();
  bag_numbering();
  set_session_time();
  calculate_classroom_population();
  generate_bulletin();
  generate_record_sheet();
  generate_small_bag_data();
  generate_big_bag_data();

  // Stop counting execution time
  newRuntime = runtime_count_stop(runtime_count_start);

  SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}


function afterRearrangeClassroomByManual(){
  let runtime_count_start = new Date();

  sort_by_session_classroom();
  bag_numbering();
  set_session_time();
  calculate_classroom_population();
  generate_bulletin();
  generate_record_sheet();
  generate_small_bag_data();
  generate_big_bag_data();

  newRuntime = runtime_count_stop(runtime_count_start);

  SpreadsheetApp.getUi().alert("已完成編排，共使用" + newRuntime + "秒");
}

function runtime_count_stop(start) {

  var stop = new Date();
  var newRuntime = Number(stop) - Number(start);
  return Math.ceil(newRuntime/1000);

}


function count_time_consume(runner){
  let start_time = new Date();
  runner()
  let end_time = new Date()
  let runtime = Math.ceil(Number(end_time) - Number(start_time))/1000;
  return runtime;
}


function set_range_values(range, data){
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
  const data_range = ss.getSheetByName("教學組排入考程的科目").getRange("A2:A");
  const data_values = data_range.getValues();
  const num_rows = data_range.getNumRows();
  const num_cols = data_range.getNumColumns();

  for(let i = 0; i < num_rows; i++) {
    if(!ss.isRowHiddenByFilter(i+1)){
      for(let j = 0; j < num_cols; j++){
        data_values[i][j] = true;
      }
    }
  }

  set_range_values(data_range, data_values);
}


function cancelCheckboxes(){
  const ss = SpreadsheetApp.getActiveSpreadsheet()
  const data_range = ss.getSheetByName("教學組排入考程的科目").getRange("A1:A");
  const data_values = data_range.getValues();
  const num_rows = data_range.getNumRows();
  const num_cols = data_range.getNumColumns();

  for(let i = 0; i < num_rows; i++) {
    if(!ss.isRowHiddenByFilter(i+1)){
      for(let j = 0; j < num_cols; j++){
        data_values[i][j] = false;
      }
    }
  }

  set_range_values(data_range, data_values);
}


function descending_population(a, b) {
    if (a[1] === b[1]) {
        return 0;
    }
    else {
        return (a[1] < b[1]) ? 1 : -1;
    }
}


function get_department_grade_statistics_of_array(data){
  let department_column = 0;
  let grade_column = 1;
  let statistics = {};
  for (row of data){
    let key = row[department_column] + row[grade_column];

    if (key in statistics){
      statistics[key] += 1;
    } else {
      statistics[key] = 1;
    }
  }

  return statistics;
}


function get_department_grade_subject_statistics_of_array(data){
  const department_column = 0;
  const grade_column = 1;
  const subject_name_column = 7;
  
  let statistics = {};
  for (row of data){
    let key = row[department_column] + row[grade_column] + "_" + row[subject_name_column];

    if (key in statistics){
      statistics[key] += 1;
    } else {
      statistics[key] = 1;
    }
  }

  return statistics;
}


function get_department_grade_statistics(){
  // 統計各科別年級、各班級的應考人數
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();

  return get_department_grade_statistics_of_array(data);
}


function get_department_grade_subject_statistics(){
  // 統計各科別年級、各班級、科目的應考人數
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();

  return get_department_grade_subject_statistics_of_array(data);
}


function create_classroom(){
  return {
    students: [],
    get population(){return this.students.length;},
    get class_subject_statisics(){
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

function create_session(){
  // session 物件工廠，用來產生下面的 get_session_statistic 函數中，需要建立 9 個 session 物件
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_CLASSROOM_NUMBER = parametersSheet.getRange("B6").getValue();
  const session = {
    classrooms: [],
    students:[], 
    get population(){return this.students.length;},

    get department_grade_statisics(){
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

    get department_class_subject_statisics(){
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
    session.classrooms.push(create_classroom())
  }
  return session;
}


function get_session_statistics(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const session_column = headers.indexOf("節次");
  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();

  const sessions = [];
  for(let i=0; i < MAX_SESSION_NUMBER + 2; i++){
    sessions.push(create_session())
  }

  for(row of data){
    sessions[row[session_column]].students.push(row);
  }

  return sessions;
}