function arrangeClassroom(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const class_column = headers.indexOf("班級");
  const subject_column = headers.indexOf("科目名稱");
  const classroom_column = headers.indexOf("試場");

  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();
  const MAX_CLASSROOM_NUMBER = parametersSheet.getRange("B6").getValue();
  const MAX_CLASSROOM_STUDENTS = parametersSheet.getRange("B7").getValue();
  const MAX_SUBJECT_NUMBER = parametersSheet.getRange("B8").getValue();

  const sessions = get_session_statistics();

  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    let students_sum = 0;  // 用來加總同節次的所有試場人數
    const dgs = Object.entries(sessions[i].department_class_subject_statisics).sort(descending_sorting);
    for(let j=1; j < sessions[i].classrooms.length; j++){
      let arranged_subjects = [];
      for(let k=0; k < dgs.length; k++){
        // 檢查此「班級-科目」是否已安排試場
        const has_duplicate = arranged_subjects.includes(dgs[k][0]);
        if(has_duplicate){
          continue;
        }

        // 檢查此試場是否還有名額
        const has_quota = dgs[k][1] + sessions[i].classrooms[j].population <= MAX_CLASSROOM_STUDENTS;
        if(!has_quota){
          continue;
        }

        // 檢查此試場的科目數是否低於限制
        const under_subject_limitation = 1 + Object.keys(sessions[i].classrooms[j].class_subject_statisics).length <= MAX_SUBJECT_NUMBER;
        if(!under_subject_limitation){
          continue;
        }

        if(sessions[i].classrooms[j].population >= MAX_CLASSROOM_STUDENTS){
          break;
        }

        if(!has_duplicate && has_quota && under_subject_limitation){
          sessions[i].students.forEach(
            function(row){
              let key=row[class_column] + row[subject_column];
              if(key == dgs[k][0] && row[classroom_column] == 0){
                row[classroom_column] = j;
                sessions[i].classrooms[j].students.push(row);
              }
            }
          );
        }

        arranged_subjects = arranged_subjects.concat(Object.keys(sessions[i].classrooms[j].class_subject_statisics));
      }
      students_sum += sessions[i].classrooms[j].population;
    }

    if(sessions[i].population != students_sum){
      let merged_session_classrooms = [];
      for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
        for(let j=1; j < sessions[i].classrooms.length; j++){
          merged_session_classrooms = merged_session_classrooms.concat(sessions[i].classrooms[j].students);
        }
      }
      break;
    }
  }

  let modified_data = [];
  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    for(let j=1; j < sessions[i].classrooms.length; j++){
      modified_data = modified_data.concat(sessions[i].classrooms[j].students);
    }
  }

  if(modified_data.length == data.length){
    set_range_values(filtered_sheet.getRange(2, 1, modified_data.length, modified_data[0].length), modified_data);
  } else {
    Logger.log("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
    SpreadsheetApp.getUi().alert("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
  }

  if (sessions[9].students.length > 0){
    SpreadsheetApp.getUi().alert("部分考生被安排在第9節補考，請注意是否需要調整到中午應試！");
  }

  filtered_sheet.getRange("I:J").setNumberFormat('#,##0');
  sort_by_session_classroom();
}
