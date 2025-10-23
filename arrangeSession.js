function arrangeCommonsSession(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const [paramHeaders,...commonData] = parametersSheet.getRange(2, 5, 21, 2).getValues();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const subjectColumn = headers.indexOf("科目名稱");
  const sessionColumn = headers.indexOf("節次");

  let commonSessions = {}
  commonData.forEach(
    function(row){
      commonSessions[row[0]] = row[1];
    }
  );
  
  let modifiedData = data.map(
    function(row){
      if(commonSessions[row[subjectColumn]] == null){
        return row;
      } else {
        row[sessionColumn] = commonSessions[row[subjectColumn]];
        return row; 
      } 
    }
  );

  if(modifiedData.length == data.length){
    setRangeValues(filteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("安排共同科目節次時，合併後的資料筆數和原有的筆數不同！");
    SpreadsheetApp.getUi().alert("安排共同科目節次時，合併後的資料筆數和原有的筆數不同！");
  }
}


function descendingSorting(a, b) {
    if (a[1] === b[1]) {
        return 0;
    }
    else {
        return (a[1] < b[1]) ? 1 : -1;
    }
}


function arrangeProfessionsSession(){
  // 安排非共同科目的節次
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const sessionColumn = headers.indexOf("節次");

  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();
  const MAX_SESSION_STUDENTS = 0.9 * parametersSheet.getRange("B9").getValue();  // 每節的最大學生數的 9 成

  const dgs = Object.entries(getDepartmentGradeSubjectStatistics()).sort(descendingSorting);
  const sessions = getSessionStatistics();
  
  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    for(let k=0; k < dgs.length; k++){
      const departmentGrade = dgs[k][0].slice(0, dgs[k][0].indexOf("_"));
      const hasDuplicate = Object.keys(sessions[i].departmentGradeStatistics).includes(departmentGrade);
      if(hasDuplicate){
        continue;
      }

      const hasQuota = dgs[k][1] + sessions[i].population <= MAX_SESSION_STUDENTS;
      if(!hasQuota){
        continue;
      }

      if(sessions[i].population >= MAX_SESSION_STUDENTS){
        Logger.log("第" + i + "節已達人數上限。");
        Logger.log("學生數為： " + sessions[i].population);
        break;
      }

      if(!hasDuplicate && hasQuota){
        data.forEach(
          function(row){
            let key=row[0] + row[1] + "_" + row[7];
            if(key == dgs[k][0] && row[8] == 0){
              row[sessionColumn] = i;
              sessions[i].students.push(row);
            }
          }
        );
      }
    }
  }

  let modifiedData = [];
  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    Logger.log("sessions[" + i + "]: " + sessions[i].population);
    modifiedData = modifiedData.concat(sessions[i].students);
  }
    
  if(modifiedData.length == data.length){
    setRangeValues(filteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("無法將所有人排入 " + MAX_SESSION_NUMBER + " 節，請檢查是否有某科年級須補考 " + parseInt(MAX_SESSION_NUMBER + 1) +" 科以上！");
    SpreadsheetApp.getUi().alert("無法將所有人排入 " + MAX_SESSION_NUMBER + "節，請檢查是否有某科年級須補考 10 科以上！");
  }
}
