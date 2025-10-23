function arrangeClassroom(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const classColumn = headers.indexOf("班級");
  const subjectColumn = headers.indexOf("科目名稱");
  const classroomColumn = headers.indexOf("試場");

  const parametersSheet = ss.getSheetByName("參數區");
  const MAX_SESSION_NUMBER = parametersSheet.getRange("B5").getValue();
  const MAX_CLASSROOM_NUMBER = parametersSheet.getRange("B6").getValue();
  const MAX_CLASSROOM_STUDENTS = parametersSheet.getRange("B7").getValue();
  const MAX_SUBJECT_NUMBER = parametersSheet.getRange("B8").getValue();

  const sessions = getSessionStatistics();

  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    let studentsSum = 0;  // 用來加總同節次的所有試場人數
    const dgs = Object.entries(sessions[i].departmentClassSubjectStatistics).sort(descendingSorting);
    for(let j=1; j < sessions[i].classrooms.length; j++){
      let arrangedSubjects = [];
      for(let k=0; k < dgs.length; k++){
        // 檢查此「班級-科目」是否已安排試場
        const hasDuplicate = arrangedSubjects.includes(dgs[k][0]);
        if(hasDuplicate){
          continue;
        }

        // 檢查此試場是否還有名額
        const hasQuota = dgs[k][1] + sessions[i].classrooms[j].population <= MAX_CLASSROOM_STUDENTS;
        if(!hasQuota){
          continue;
        }

        // 檢查此試場的科目數是否低於限制
        const underSubjectLimitation = 1 + Object.keys(sessions[i].classrooms[j].classSubjectStatistics).length <= MAX_SUBJECT_NUMBER;
        if(!underSubjectLimitation){
          continue;
        }

        if(sessions[i].classrooms[j].population >= MAX_CLASSROOM_STUDENTS){
          break;
        }

        if(!hasDuplicate && hasQuota && underSubjectLimitation){
          sessions[i].students.forEach(
            function(row){
              let key=row[classColumn] + row[subjectColumn];
              if(key == dgs[k][0] && row[classroomColumn] == 0){
                row[classroomColumn] = j;
                sessions[i].classrooms[j].students.push(row);
              }
            }
          );
        }

        arrangedSubjects = arrangedSubjects.concat(Object.keys(sessions[i].classrooms[j].classSubjectStatistics));
      }
      studentsSum += sessions[i].classrooms[j].population;
    }

    if(sessions[i].population != studentsSum){
      let mergedSessionClassrooms = [];
      for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
        for(let j=1; j < sessions[i].classrooms.length; j++){
          mergedSessionClassrooms = mergedSessionClassrooms.concat(sessions[i].classrooms[j].students);
        }
      }
      break;
    }
  }

  let modifiedData = [];
  for(let i=1; i < MAX_SESSION_NUMBER + 2; i++){
    for(let j=1; j < sessions[i].classrooms.length; j++){
      modifiedData = modifiedData.concat(sessions[i].classrooms[j].students);
    }
  }

  if(modifiedData.length == data.length){
    setRangeValues(filteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
    SpreadsheetApp.getUi().alert("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
  }

  if (sessions[9].students.length > 0){
    SpreadsheetApp.getUi().alert("部分考生被安排在第9節補考，請注意是否需要調整到中午應試！");
  }

  filteredSheet.getRange("I:J").setNumberFormat('#,##0');
  sortBySessionClassroom();
}
