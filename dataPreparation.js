function updateUnfilteredSubjectCodes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const unfilteredSource = spreadsheet.getSheetByName("註冊組補考名單");
  const [unfilteredHeaders, ...unfilteredRows] = unfilteredSource.getDataRange().getValues();

  const classColumnIndex = unfilteredHeaders.indexOf("班級");
  const subjectColumnIndex = unfilteredHeaders.indexOf("科目");
  const codeColumnIndex = unfilteredHeaders.indexOf("科目代碼補完");
  const subjectNameColumnIndex = unfilteredHeaders.indexOf("科目名稱");

  const departmentGroupMap = {
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

  const gradeYearMap = {
    "一": parseInt(parameterSheet.getRange("B2").getValue()),
    "二": parseInt(parameterSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parameterSheet.getRange("B2").getValue()) - 2,
  };

  const patchedRows = [];
  unfilteredRows.forEach(
    function(candidateRow){
      const rawSubjectCell = candidateRow[subjectColumnIndex].toString();
      const rawSubjectCode = rawSubjectCell.split(".")[0];
      if(rawSubjectCode.length === 16){
        candidateRow[codeColumnIndex] = rawSubjectCode.slice(0, 3) + "553401" + rawSubjectCode.slice(3, 9) + "0" + rawSubjectCode.slice(9);
      } else {
        const gradeDigit = candidateRow[classColumnIndex].toString().slice(2, 3);
        const departmentDigits = rawSubjectCode.slice(0,3);
        candidateRow[codeColumnIndex] = gradeYearMap[gradeDigit] + "553401V" + departmentGroupMap[departmentDigits] + departmentDigits + "0" + rawSubjectCode.slice(3);
      }

      candidateRow[subjectNameColumnIndex] = rawSubjectCell.split(".")[1];
      patchedRows.push(candidateRow);
    }
  );

  if(patchedRows.length === unfilteredRows.length){
    writeRangeValuesSafely(unfilteredSource.getRange(2, 1, patchedRows.length, patchedRows[0].length), patchedRows);
  } else {
    Logger.log("課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("課程代碼補完失敗！");
  }
}


function updateOpenCourseCodes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const openCourseSheet = spreadsheet.getSheetByName("開課資料(查詢任課教師用)");
  const [openHeaders, ...openRows] = openCourseSheet.getDataRange().getValues();

  const classColumnIndex = openHeaders.indexOf("班級名稱");
  const codeColumnIndex = openHeaders.indexOf("科目代碼");
  const completedCodeColumnIndex = openHeaders.indexOf("科目代碼補完");

  const departmentGroupMap = {
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

  const gradeYearMap = {
    "一": parseInt(parameterSheet.getRange("B2").getValue()),
    "二": parseInt(parameterSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parameterSheet.getRange("B2").getValue()) - 2,
  };

  const patchedRows = [];
  openRows.forEach(
    function(courseRow){
      const rawSubjectCode = courseRow[codeColumnIndex];
      if(rawSubjectCode.length === 16){
        courseRow[completedCodeColumnIndex] = rawSubjectCode.slice(0, 3) + "553401" + rawSubjectCode.slice(3, 9) + "0" + rawSubjectCode.slice(9);
      } else {
        const gradeDigit = courseRow[classColumnIndex].toString().slice(2, 3);
        const departmentDigits = rawSubjectCode.slice(0,3);
        courseRow[completedCodeColumnIndex] = gradeYearMap[gradeDigit] + "553401V" + departmentGroupMap[departmentDigits] + departmentDigits + "0" + rawSubjectCode.slice(3);
      }
      patchedRows.push(courseRow);
    }
  );

  if(patchedRows.length === openRows.length){
    writeRangeValuesSafely(openCourseSheet.getRange(2, 1, patchedRows.length, patchedRows[0].length), patchedRows);
  } else {
    Logger.log("開課資料課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("開課資料課程代碼補完失敗！");
  }
}


function buildFilteredCandidateList(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const unfilteredSource = spreadsheet.getSheetByName("註冊組補考名單");
  const candidateSheet = spreadsheet.getSheetByName("教學組排入考程的科目");
  const openCourseSheet = spreadsheet.getSheetByName("開課資料(查詢任課教師用)");
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [candidateHeaders, ...candidateRows] = candidateSheet.getDataRange().getValues();
  const [unfilteredHeaders, ...unfilteredRows] = unfilteredSource.getDataRange().getValues();
  const [openHeaders, ...openRows] = openCourseSheet.getDataRange().getValues();

  const studentNumberIndex = unfilteredHeaders.indexOf("學號");
  const classNameIndex = unfilteredHeaders.indexOf("班級");
  const seatNumberIndex = unfilteredHeaders.indexOf("座號");
  const studentNameIndex = unfilteredHeaders.indexOf("姓名");
  const subjectNameIndex = unfilteredHeaders.indexOf("科目名稱");
  const completedCodeIndex = unfilteredHeaders.indexOf("科目代碼補完");
  const openClassIndex = openHeaders.indexOf("班級名稱");
  const openSubjectNameIndex = openHeaders.indexOf("科目名稱");
  const teacherNameIndex = openHeaders.indexOf("任課教師");

  const makeUpRequiredIndex = candidateHeaders.indexOf("要補考");
  const filteredCodeIndex = candidateHeaders.indexOf("課程代碼");
  const requiresComputerIndex = candidateHeaders.indexOf("電腦");
  const requiresManualIndex = candidateHeaders.indexOf("人工");

  const eligibleSubjectsByCode = {};
  candidateRows.forEach(
    function(candidateRow){
      if(candidateRow[makeUpRequiredIndex] === true){
        eligibleSubjectsByCode[candidateRow[filteredCodeIndex]] = {
          computerRequired: candidateRow[requiresComputerIndex],
          manualRequired: candidateRow[requiresManualIndex],
        };
      }
    }
  );

  const teacherLookup = {};
  openRows.forEach(
    function(courseRow){
      const teacherCell = courseRow[teacherNameIndex].toString();
      const lookupKey = courseRow[openClassIndex].toString() + courseRow[openSubjectNameIndex].toString();
      if (teacherCell.length > 10){
        teacherLookup[lookupKey] = teacherCell.split(",")[0].slice(7);
      } else {
        teacherLookup[lookupKey] = teacherCell.slice(6);
      }
    }
  );

  // 清除「排入考程的補考名單」內容
  resetFilteredSheets();

  const filteredRows = [];
  unfilteredRows.forEach(
    function(studentRow){
      if (Object.keys(eligibleSubjectsByCode).includes(studentRow[completedCodeIndex])){
        const courseKey = studentRow[classNameIndex].toString() + studentRow[subjectNameIndex].toString();
        const computerMark = eligibleSubjectsByCode[studentRow[completedCodeIndex]].computerRequired ? "☑" : "☐";
        const manualMark = eligibleSubjectsByCode[studentRow[completedCodeIndex]].manualRequired ? "☑" : "☐";
        const filteredRow = [
          lookupDepartmentName(studentRow[classNameIndex]), // 科別
          deriveGradeLevel(studentRow[classNameIndex]), // 年級
          deriveClassCode(studentRow[classNameIndex]), // 班級代碼
          studentRow[classNameIndex], // 班級
          studentRow[seatNumberIndex], // 座號
          studentRow[studentNumberIndex], // 學號
          studentRow[studentNameIndex], // 姓名
          studentRow[subjectNameIndex], // 科目名稱
          studentRow[8] = 0,  // 節次預設為0
          studentRow[9] = 0,  // 試場預設為0
          "",  // 小袋序號
          "",  // 小袋人數
          "",  // 大袋序號
          "",  // 大袋人數
          "",  // 班級人數
          "",  // 時間
          computerMark,
          manualMark,
          teacherLookup[courseKey]  // 任課老師
        ];

        filteredRows.push(filteredRow);
      }
    }
  );

  filteredSheet.getRange(2, 1, filteredRows.length, filteredRows[0].length)
    .setNumberFormat('@STRING@')  // 改成純文字格式，以免 0 開頭的學號被去掉前面的 0，造成位數錯誤
    .setValues(filteredRows);
}


function markVisibleCandidateCheckboxes(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const candidateRange = spreadsheet.getSheetByName("教學組排入考程的科目").getRange("A2:A");
  const checkboxValues = candidateRange.getValues();
  const totalRows = candidateRange.getNumRows();
  const totalColumns = candidateRange.getNumColumns();

  for(let rowIndex = 0; rowIndex < totalRows; rowIndex++){
    if(!spreadsheet.isRowHiddenByFilter(rowIndex + 1)){
      for(let columnIndex = 0; columnIndex < totalColumns; columnIndex++){
        checkboxValues[rowIndex][columnIndex] = true;
      }
    }
  }

  writeRangeValuesSafely(candidateRange, checkboxValues);
}


function clearCandidateCheckboxes(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const candidateRange = spreadsheet.getSheetByName("教學組排入考程的科目").getRange("A1:A");
  const checkboxValues = candidateRange.getValues();
  const totalRows = candidateRange.getNumRows();
  const totalColumns = candidateRange.getNumColumns();

  for(let rowIndex = 0; rowIndex < totalRows; rowIndex++){
    if(!spreadsheet.isRowHiddenByFilter(rowIndex + 1)){
      for(let columnIndex = 0; columnIndex < totalColumns; columnIndex++){
        checkboxValues[rowIndex][columnIndex] = false;
      }
    }
  }

  writeRangeValuesSafely(candidateRange, checkboxValues);
}
