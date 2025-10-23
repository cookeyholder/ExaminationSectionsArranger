function arrange_commons_session(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parameters_sheet = ss.getSheetByName("參數區");
  const [param_headers,...commonData] = parameters_sheet.getRange(2, 5, 21, 2).getValues();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const subject_column = headers.indexOf("科目名稱");
  const session_column = headers.indexOf("節次");

  let commonSessions = {}
  commonData.forEach(
    function(row){
      commonSessions[row[0]] = row[1];
    }
  );
  
  let modified_data = data.map(
    function(row){
      if(commonSessions[row[subject_column]] == null){
        return row;
      } else {
        row[session_column] = commonSessions[row[subject_column]];
        return row; 
      } 
    }
  );

  if(modified_data.length == data.length){
    set_range_values(filtered_sheet.getRange(2, 1, modified_data.length, modified_data[0].length), modified_data);
  } else {
    Logger.log("安排共同科目節次時，合併後的資料筆數和原有的筆數不同！");
    SpreadsheetApp.getUi().alert("安排共同科目節次時，合併後的資料筆數和原有的筆數不同！");
  }
}


function descending_sorting(a, b) {
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
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const session_column = headers.indexOf("節次");

  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();
  const MAX_SESSION_STUDENTS = 0.9 * parametersSheet.getRange("B9").getValue();  // 每節的最大學生數的 9 成

  const dgs = Object.entries(get_department_grade_subject_statistics()).sort(descending_sorting);
  const sessions = get_session_statistics();
  
  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    for(let k=0; k < dgs.length; k++){
      const department_grade = dgs[k][0].slice(0, dgs[k][0].indexOf("_"));
      const has_duplicate = Object.keys(sessions[i].department_grade_statisics).includes(department_grade);
      if(has_duplicate){
        continue;
      }

      const has_quota = dgs[k][1] + sessions[i].population <= MAX_SESSION_STUDENTS;
      if(!has_quota){
        continue;
      }

      if(sessions[i].population >= MAX_SESSION_STUDENTS){
        Logger.log("第" + i + "節已達人數上限。");
        Logger.log("學生數為： " + sessions[i].population);
        break;
      }

      if(!has_duplicate && has_quota){
        data.forEach(
          function(row){
            let key=row[0] + row[1] + "_" + row[7];
            if(key == dgs[k][0] && row[8] == 0){
              row[session_column] = i;
              sessions[i].students.push(row);
            }
          }
        );
      }
    }
  }

  let modified_data = [];
  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    Logger.log("sessions[" + i + "]: " + sessions[i].population);
    modified_data = modified_data.concat(sessions[i].students);
  }
    
  if(modified_data.length == data.length){
    set_range_values(filtered_sheet.getRange(2, 1, modified_data.length, modified_data[0].length), modified_data);
  } else {
    Logger.log("無法將所有人排入 " + MAX_SESSION_NUMBER + " 節，請檢查是否有某科年級須補考 " + parseInt(MAX_SESSION_NUMBER + 1) +" 科以上！");
    SpreadsheetApp.getUi().alert("無法將所有人排入 " + MAX_SESSION_NUMBER + "節，請檢查是否有某科年級須補考 10 科以上！");
  }
}
