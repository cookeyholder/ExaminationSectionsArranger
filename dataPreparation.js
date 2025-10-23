function unfilteredCodeComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const unfilteredSheet = ss.getSheetByName("註冊組補考名單");
  const [unfilteredSheetHeaders, ...unfilteredData] = unfilteredSheet.getDataRange().getValues();

  const classColumn = unfilteredSheetHeaders.indexOf("班級");
  const subjectColumn = unfilteredSheetHeaders.indexOf("科目");
  const codeColumn = unfilteredSheetHeaders.indexOf("科目代碼補完");
  const subjectNameColumn = unfilteredSheetHeaders.indexOf("科目名稱");

  const departmentToGroup = {
    "301": "21",
    "303": "22",
    "305": "23",
    "306": "23",
    "308": "23",
    "309": "23",
    "311": "25",
    "315": "24",
    "373": "28",
    "374": "21",
  };

  const gradeToYear = {
    "一": parseInt(parametersSheet.getRange("B2").getValue()),
    "二": parseInt(parametersSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parametersSheet.getRange("B2").getValue()) - 2,
  };

  let modifiedData = [];
  unfilteredData.forEach(
    function(row){
      let tmp = row[subjectColumn].toString().split(".")[0];
      if(tmp.length == 16){
        row[codeColumn] = tmp.slice(0, 3) + "553401" + tmp.slice(3, 9) + "0" + tmp.slice(9);
      } else {
        row[codeColumn] = gradeToYear[row[classColumn].toString().slice(2, 3)] + "553401V" + departmentToGroup[tmp.slice(0,3)] + tmp.slice(0,3) + "0" + tmp.slice(3);
      }

      row[subjectNameColumn] = row[subjectColumn].toString().split(".")[1];
      modifiedData.push(row);
    }
  );

  if(modifiedData.length == unfilteredData.length){
    setRangeValues(unfilteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("課程代碼補完失敗！");
  }
}



function openCodeComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const openSheet = ss.getSheetByName("開課資料(查詢任課教師用)");
  const [openSheetHeaders, ...openData] = openSheet.getDataRange().getValues();

  const classColumn = openSheetHeaders.indexOf("班級名稱");
  const codeColumn = openSheetHeaders.indexOf("科目代碼");
  const completeColumn = openSheetHeaders.indexOf("科目代碼補完");
  const subjectNameColumn = openSheetHeaders.indexOf("科目名稱");

  const departmentToGroup = {
    "301": "21",
    "303": "22",
    "305": "23",
    "306": "23",
    "308": "23",
    "309": "23",
    "311": "25",
    "315": "24",
    "373": "28",
    "374": "21",
  };

  const gradeToYear = {
    "一": parseInt(parametersSheet.getRange("B2").getValue()),
    "二": parseInt(parametersSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parametersSheet.getRange("B2").getValue()) - 2,
  };

  let modifiedData = [];
  openData.forEach(
    function(row){
      let tmp = row[codeColumn];
      if(row[codeColumn].length == 16){
        row[completeColumn] = tmp.slice(0, 3) + "553401" + tmp.slice(3, 9) + "0" + tmp.slice(9);
      } else {
        row[completeColumn] = gradeToYear[row[classColumn].toString().slice(2, 3)] + "553401V" + departmentToGroup[tmp.slice(0,3)] + tmp.slice(0,3) + "0" + tmp.slice(3);
      }


      modifiedData.push(row);
    }
  );

  if(modifiedData.length == openData.length){
    setRangeValues(openSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("開課資料課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("開課資料課程代碼補完失敗！");
  }
}


function getFilteredData(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unfilteredSheet = ss.getSheetByName("註冊組補考名單");
  const candidateSheet = ss.getSheetByName("教學組排入考程的科目");
  const openSheet = ss.getSheetByName("開課資料(查詢任課教師用)");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [candidateSubjectHeaders, ...candidateSubjectsData] = candidateSheet.getDataRange().getValues();
  const [unfilteredSheetHeaders, ...unfilteredData] = unfilteredSheet.getDataRange().getValues();
  const [openSheetHeaders, ...openData] = openSheet.getDataRange().getValues();

  const stdNumberColumn = unfilteredSheetHeaders.indexOf("學號");
  const classColumn = unfilteredSheetHeaders.indexOf("班級");
  const seatNumberColumn = unfilteredSheetHeaders.indexOf("座號");
  const stdNameColumn = unfilteredSheetHeaders.indexOf("姓名");
  const subjectNameColumn = unfilteredSheetHeaders.indexOf("科目名稱");
  const codeColumn = unfilteredSheetHeaders.indexOf("科目代碼補完");
  const openClassColumn = openSheetHeaders.indexOf("班級名稱");
  const openSubjectNameColumn = openSheetHeaders.indexOf("科目名稱");
  const teacherColumn = openSheetHeaders.indexOf("任課教師");

  const makeUpColumn = candidateSubjectHeaders.indexOf("要補考");
  const filteredCodeColumn = candidateSubjectHeaders.indexOf("課程代碼");
  const byComputerColumn = candidateSubjectHeaders.indexOf("電腦");
  const byHandColumn = candidateSubjectHeaders.indexOf("人工");

  let candidateSubjects = {};
  candidateSubjectsData.forEach(
    function (row){
      if(row[makeUpColumn]==true){
        candidateSubjects[row[filteredCodeColumn]] = {
          "bycomputer": row[byComputerColumn],
          "byhand": row[byHandColumn]
        };
      }
    }
  );

  let openTeacher = {};
  openData.forEach(
    function (row){
      if (row[teacherColumn].toString().length > 10){
        openTeacher[row[openClassColumn].toString() + row[openSubjectNameColumn].toString()] = row[teacherColumn].toString().split(",")[0].slice(7);
      } else {
        openTeacher[row[openClassColumn].toString() + row[openSubjectNameColumn].toString()] = row[teacherColumn].toString().slice(6);
      }
    }
  );

  // 清除「排入考程的補考名單」內容
  initialize();

  let nameList = [];
  unfilteredData.forEach(
    function(row){
      if (Object.keys(candidateSubjects).includes(row[codeColumn])){
        // 科別	年級	班級代碼	班級	座號	學號	姓名	科目名稱	節次	試場	小袋序號	小袋人數	大袋序號	大袋人數	班級人數	時間	電腦	人工	任課老師
        let tmp = [
          getDepartmentName(row[classColumn]), // 科別
          getGrade(row[classColumn]), // 年級
          getClassCode(row[classColumn]), // 班級代碼
          row[classColumn], // 班級
          row[seatNumberColumn], // 座號
          row[stdNumberColumn], // 學號
          row[stdNameColumn], // 姓名
          row[subjectNameColumn], // 科目名稱
          row[8]=0,  // 節次預設為0
          row[9]=0,  // 試場預設為0
          "",  // 小袋序號
          "",  // 小袋人數
          "",  // 大袋序號
          "",  // 大袋人數
          "",  // 班級人數
          "",  // 時間
          candidateSubjects[row[codeColumn]]["bycomputer"] ? "☑" : "☐",  // 電腦
          candidateSubjects[row[codeColumn]]["byhand"] ? "☑" : "☐",  //人工
          openTeacher[row[classColumn].toString() + row[subjectNameColumn].toString()]  // 任課老師
        ];

        nameList.push(tmp);
      }
    }
  );

  filteredSheet.getRange(2, 1, nameList.length, nameList[0].length)
    .setNumberFormat('@STRING@')  // 改成純文字格式，以免 0 開頭的學號被去掉前面的 0，造成位數錯誤
    .setValues(nameList);
}


function checkShowedBoxes(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
