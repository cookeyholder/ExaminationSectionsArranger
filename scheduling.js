function scheduleCommonSubjectSessions(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const sessionRuleRows = parameterSheet.getRange(2, 5, 21, 2).getValues();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const subjectNameIndex = headerRow.indexOf("科目名稱");
  const sessionIndex = headerRow.indexOf("節次");

  const preferredSessionBySubject = {};
  sessionRuleRows.forEach(
    function(ruleRow){
      if (ruleRow[0] && ruleRow[1]){
        preferredSessionBySubject[ruleRow[0]] = ruleRow[1];
      }
    }
  );

  const updatedRows = candidateRows.map(
    function(examineeRow){
      const preferredSession = preferredSessionBySubject[examineeRow[subjectNameIndex]];
      if (preferredSession == null){
        return examineeRow;
      }
      examineeRow[sessionIndex] = preferredSession;
      return examineeRow;
    }
  );

  if(updatedRows.length === candidateRows.length){
    writeRangeValuesSafely(filteredSheet.getRange(2, 1, updatedRows.length, updatedRows[0].length), updatedRows);
  } else {
    Logger.log("安排共同科目節次時，合併後的資料筆數和原有的筆數不同！");
    SpreadsheetApp.getUi().alert("安排共同科目節次時，合併後的資料筆數和原有的筆數不同！");
  }
}


function scheduleSpecializedSubjectSessions(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const sessionIndex = headerRow.indexOf("節次");

  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const maxSessionCount = parameterSheet.getRange("B5").getValue();
  const sessionCapacity = 0.9 * parameterSheet.getRange("B9").getValue();

  const departmentGradeSubjectCounts = Object.entries(fetchDepartmentGradeSubjectCounts()).sort(compareCountDescending);
  const sessionSnapshots = buildSessionStatistics();

  for(let sessionNumber = 1; sessionNumber < maxSessionCount + 2; sessionNumber++){
    for(let countIndex = 0; countIndex < departmentGradeSubjectCounts.length; countIndex++){
      const departmentGradeKey = departmentGradeSubjectCounts[countIndex][0].slice(0, departmentGradeSubjectCounts[countIndex][0].indexOf("_"));
      const isDepartmentScheduled = Object.keys(sessionSnapshots[sessionNumber].departmentGradeStatistics).includes(departmentGradeKey);
      if(isDepartmentScheduled){
        continue;
      }

      const hasCapacity = departmentGradeSubjectCounts[countIndex][1] + sessionSnapshots[sessionNumber].population <= sessionCapacity;
      if(!hasCapacity){
        continue;
      }

      if(sessionSnapshots[sessionNumber].population >= sessionCapacity){
        Logger.log("第" + sessionNumber + "節已達人數上限。");
        Logger.log("學生數為： " + sessionSnapshots[sessionNumber].population);
        break;
      }

      candidateRows.forEach(
        function(examineeRow){
          const departmentGradeSubjectKey = examineeRow[0] + examineeRow[1] + "_" + examineeRow[7];
          if(departmentGradeSubjectKey === departmentGradeSubjectCounts[countIndex][0] && examineeRow[8] === 0){
            examineeRow[sessionIndex] = sessionNumber;
            sessionSnapshots[sessionNumber].students.push(examineeRow);
          }
        }
      );
    }
  }

  let combinedRows = [];
  for(let sessionNumber = 1; sessionNumber < maxSessionCount + 2; sessionNumber++){
    Logger.log("sessions[" + sessionNumber + "]: " + sessionSnapshots[sessionNumber].population);
    combinedRows = combinedRows.concat(sessionSnapshots[sessionNumber].students);
  }
    
  if(combinedRows.length === candidateRows.length){
    writeRangeValuesSafely(filteredSheet.getRange(2, 1, combinedRows.length, combinedRows[0].length), combinedRows);
  } else {
    Logger.log("無法將所有人排入 " + maxSessionCount + " 節，請檢查是否有某科年級須補考 " + parseInt(maxSessionCount + 1) +" 科以上！");
    SpreadsheetApp.getUi().alert("無法將所有人排入 " + maxSessionCount + "節，請檢查是否有某科年級須補考 10 科以上！");
  }
}


function assignExamRooms(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const classIndex = headerRow.indexOf("班級");
  const subjectIndex = headerRow.indexOf("科目名稱");
  const roomIndex = headerRow.indexOf("試場");

  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const maxSessionCount = parameterSheet.getRange("B5").getValue();
  const maxRoomCount = parameterSheet.getRange("B6").getValue();
  const maxStudentsPerRoom = parameterSheet.getRange("B7").getValue();
  const maxSubjectsPerRoom = parameterSheet.getRange("B8").getValue();

  const sessionSnapshots = buildSessionStatistics();

  for(let sessionNumber = 1; sessionNumber < maxSessionCount + 2; sessionNumber++){
    let totalStudentsWithinSession = 0;
    const departmentSubjectCounts = Object.entries(sessionSnapshots[sessionNumber].departmentClassSubjectStatistics).sort(compareCountDescending);
    for(let roomNumber = 1; roomNumber < sessionSnapshots[sessionNumber].classrooms.length; roomNumber++){
      let scheduledSubjects = [];
      for(let countIndex = 0; countIndex < departmentSubjectCounts.length; countIndex++){
        const isSubjectScheduled = scheduledSubjects.includes(departmentSubjectCounts[countIndex][0]);
        if(isSubjectScheduled){
          continue;
        }

        const roomHasCapacity = departmentSubjectCounts[countIndex][1] + sessionSnapshots[sessionNumber].classrooms[roomNumber].population <= maxStudentsPerRoom;
        if(!roomHasCapacity){
          continue;
        }

        const subjectCountWithinRoom = Object.keys(sessionSnapshots[sessionNumber].classrooms[roomNumber].classSubjectStatistics).length;
        const underSubjectLimit = 1 + subjectCountWithinRoom <= maxSubjectsPerRoom;
        if(!underSubjectLimit){
          continue;
        }

        if(sessionSnapshots[sessionNumber].classrooms[roomNumber].population >= maxStudentsPerRoom){
          break;
        }

        sessionSnapshots[sessionNumber].students.forEach(
          function(examineeRow){
            const subjectKey = examineeRow[classIndex] + examineeRow[subjectIndex];
            if(subjectKey === departmentSubjectCounts[countIndex][0] && examineeRow[roomIndex] === 0){
              examineeRow[roomIndex] = roomNumber;
              sessionSnapshots[sessionNumber].classrooms[roomNumber].students.push(examineeRow);
            }
          }
        );

        scheduledSubjects = scheduledSubjects.concat(Object.keys(sessionSnapshots[sessionNumber].classrooms[roomNumber].classSubjectStatistics));
      }
      totalStudentsWithinSession += sessionSnapshots[sessionNumber].classrooms[roomNumber].population;
    }

    if(sessionSnapshots[sessionNumber].population !== totalStudentsWithinSession){
      break;
    }
  }

  let reorderedRows = [];
  for(let sessionNumber = 1; sessionNumber < maxSessionCount + 2; sessionNumber++){
    for(let roomNumber = 1; roomNumber < sessionSnapshots[sessionNumber].classrooms.length; roomNumber++){
      reorderedRows = reorderedRows.concat(sessionSnapshots[sessionNumber].classrooms[roomNumber].students);
    }
  }

  if(reorderedRows.length === candidateRows.length){
    writeRangeValuesSafely(filteredSheet.getRange(2, 1, reorderedRows.length, reorderedRows[0].length), reorderedRows);
  } else {
    Logger.log("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
    SpreadsheetApp.getUi().alert("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
  }

  const sessionSnapshotsForAlert = buildSessionStatistics();
  if (sessionSnapshotsForAlert[9].students.length > 0){
    SpreadsheetApp.getUi().alert("部分考生被安排在第9節補考，請注意是否需要調整到中午應試！");
  }

  filteredSheet.getRange("I:J").setNumberFormat('#,##0');
  sortFilteredStudentsBySessionRoom();
}


function allocateBagIdentifiers() {
  sortFilteredStudentsBySessionRoom();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const classIndex = headerRow.indexOf("班級");
  const subjectIndex = headerRow.indexOf("科目名稱");
  const sessionIndex = headerRow.indexOf("節次");
  const roomIndex = headerRow.indexOf("試場");
  const smallBagIndex = headerRow.indexOf("小袋序號");
  const bigBagIndex = headerRow.indexOf("大袋序號");

  const smallBagLookup = {};
  const bigBagLookup = {};
  let nextSmallBagNumber = 0;
  let nextBigBagNumber = 0;

  candidateRows.forEach(
    function(examineeRow){
      const bigBagKey = examineeRow[sessionIndex] + "-" + examineeRow[roomIndex];
      const smallBagKey = examineeRow[sessionIndex] + "-" + examineeRow[roomIndex] + "_" + examineeRow[classIndex] + "=" + examineeRow[subjectIndex];

      if(!Object.keys(bigBagLookup).includes(bigBagKey)){
        bigBagLookup[bigBagKey] = nextBigBagNumber + 1;
        nextBigBagNumber += 1;
      }

      if(!Object.keys(smallBagLookup).includes(smallBagKey)){
        smallBagLookup[smallBagKey] = nextSmallBagNumber + 1;
        nextSmallBagNumber += 1;
      }
    }
  );

  candidateRows.forEach(
    function(examineeRow){
      const bigBagKey = examineeRow[sessionIndex] + "-" + examineeRow[roomIndex];
      const smallBagKey = examineeRow[sessionIndex] + "-" + examineeRow[roomIndex] + "_" + examineeRow[classIndex] + "=" + examineeRow[subjectIndex];
      examineeRow[bigBagIndex] = bigBagLookup[bigBagKey];
      examineeRow[smallBagIndex] = smallBagLookup[smallBagKey];
    }
  );

  writeRangeValuesSafely(filteredSheet.getRange(2, 1, candidateRows.length, candidateRows[0].length), candidateRows);
}


function populateSessionTimes() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const sessionIndex = headerRow.indexOf("節次");
  const timeIndex = headerRow.indexOf("時間");

  const sessionTimeSheet = spreadsheet.getSheetByName("節次時間表");
  const [timeHeaders, ...sessionTimeRows] = sessionTimeSheet.getDataRange().getValues();

  const sessionTimeLookup = {};
  sessionTimeRows.forEach(
    function(timeRow){
      sessionTimeLookup[timeRow[0]] = timeRow[1];
    }
  );

  const updatedRows = candidateRows.map(
    function(examineeRow){
      examineeRow[timeIndex] = sessionTimeLookup[examineeRow[sessionIndex]];
      return examineeRow;
    }
  );

  if(updatedRows.length === candidateRows.length){
    writeRangeValuesSafely(filteredSheet.getRange(2, 1, updatedRows.length, updatedRows[0].length), updatedRows);
  } else {
    Logger.log("寫入節次時間失敗！");
    SpreadsheetApp.getUi().alert("寫入節次時間失敗！");
  }
}


function updateBagAndClassPopulations() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const classIndex = headerRow.indexOf("班級");
  const bigBagIndex = headerRow.indexOf("大袋序號");
  const smallBagIndex = headerRow.indexOf("小袋序號");
  const bigBagPopulationIndex = headerRow.indexOf("大袋人數");
  const smallBagPopulationIndex = headerRow.indexOf("小袋人數");
  const classPopulationIndex = headerRow.indexOf("班級人數");

  const bigBagCounts = {};
  const smallBagCounts = {};
  const classCounts = {};
  candidateRows.forEach(
    function(examineeRow){
      const bigBagKey = "大袋" + examineeRow[bigBagIndex];
      if(Object.keys(bigBagCounts).includes(bigBagKey)){
        bigBagCounts[bigBagKey] += 1;
      } else {
        bigBagCounts[bigBagKey] = 1;
      }

      const smallBagKey = "小袋" + examineeRow[smallBagIndex];
      if(Object.keys(smallBagCounts).includes(smallBagKey)){
        smallBagCounts[smallBagKey] += 1;
      } else {
        smallBagCounts[smallBagKey] = 1;
      }

      const classKey = "班級" + examineeRow[smallBagIndex] + examineeRow[classIndex];
      if(Object.keys(classCounts).includes(classKey)){
        classCounts[classKey] += 1;
      } else {
        classCounts[classKey] = 1;
      }
    }
  );

  candidateRows.forEach(
    function(examineeRow){
      const bigBagKey = "大袋" + examineeRow[bigBagIndex];
      const smallBagKey = "小袋" + examineeRow[smallBagIndex];
      const classKey = "班級" + examineeRow[smallBagIndex] + examineeRow[classIndex];
      examineeRow[bigBagPopulationIndex] = bigBagCounts[bigBagKey];
      examineeRow[smallBagPopulationIndex] = smallBagCounts[smallBagKey];
      examineeRow[classPopulationIndex] = classCounts[classKey];
    }
  );

  writeRangeValuesSafely(filteredSheet.getRange(2, 1, candidateRows.length, candidateRows[0].length), candidateRows);
}


function sortFilteredStudentsBySubject(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const filteredRange = filteredSheet.getDataRange();

  filteredRange.offset(1,0,filteredRange.getNumRows()-1).sort(
    [
      {column: 2, ascending: true}, 
      {column: 3, ascending: true},
      {column: 9, ascending: true},
      {column: 10, ascending: true}, 
      {column: 8, ascending: true}, 
      {column: 5, ascending: true}      
    ]
  );
}


function sortFilteredStudentsByClassSeat(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const filteredRange = filteredSheet.getDataRange();

  filteredRange.offset(1,0,filteredRange.getNumRows()-1).sort(
    [
      {column: 2, ascending: true}, 
      {column: 3, ascending: true}, 
      {column: 5, ascending: true}, 
      {column: 9, ascending: true}, 
      {column: 8, ascending: true}
    ]
  );
}


function sortFilteredStudentsBySessionRoom(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const filteredRange = filteredSheet.getDataRange();

  filteredRange.offset(1,0,filteredRange.getNumRows()-1).sort(
    [
      {column: 9, ascending: true},
      {column: 10, ascending: true},
      {column: 2, ascending: true},
      {column: 3, ascending: true},
      {column: 5, ascending: true},
      {column: 8, ascending: true},
    ]
  );
}


function compareCountDescending(firstEntry, secondEntry) {
  if (firstEntry[1] === secondEntry[1]) {
    return 0;
  }
  return (firstEntry[1] < secondEntry[1]) ? 1 : -1;
}


function summarizeDepartmentGradeCounts(rowData){
  const departmentIndex = 0;
  const gradeIndex = 1;
  const aggregateCounts = {};
  for (const row of rowData){
    const key = row[departmentIndex] + row[gradeIndex];

    if (key in aggregateCounts){
      aggregateCounts[key] += 1;
    } else {
      aggregateCounts[key] = 1;
    }
  }

  return aggregateCounts;
}


function summarizeDepartmentGradeSubjectCounts(rowData){
  const departmentIndex = 0;
  const gradeIndex = 1;
  const subjectNameIndex = 7;
  
  const aggregateCounts = {};
  for (const row of rowData){
    const key = row[departmentIndex] + row[gradeIndex] + "_" + row[subjectNameIndex];

    if (key in aggregateCounts){
      aggregateCounts[key] += 1;
    } else {
      aggregateCounts[key] = 1;
    }
  }

  return aggregateCounts;
}


function fetchDepartmentGradeCounts(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [, ...candidateRows] = filteredSheet.getDataRange().getValues();

  return summarizeDepartmentGradeCounts(candidateRows);
}


function fetchDepartmentGradeSubjectCounts(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [, ...candidateRows] = filteredSheet.getDataRange().getValues();

  return summarizeDepartmentGradeSubjectCounts(candidateRows);
}


function createEmptyClassroomRecord(){
  return {
    students: [],
    get population(){return this.students.length;},
    get classSubjectStatistics(){
      const statistics = {};
      this.students.forEach(
        function(studentRow){
          const key = studentRow[3] + "_" + studentRow[7];  // 班級 + _ + 科目
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
}


function createEmptySessionRecord(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const maxRoomCount = parameterSheet.getRange("B6").getValue();
  const sessionRecord = {
    classrooms: [],
    students:[], 
    get population(){return this.students.length;},

    get departmentGradeStatistics(){
      const statistics = {};
      this.students.forEach(
        function(studentRow){
          const key = studentRow[0] + studentRow[1];  // 科別 + 年級
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
      const statistics = {};
      this.students.forEach(
        function(studentRow){
          const key = studentRow[3] + studentRow[7];  // 班級 + 科目
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

  for(let roomIndex = 0; roomIndex < maxRoomCount + 1; roomIndex++){
    sessionRecord.classrooms.push(createEmptyClassroomRecord());
  }
  return sessionRecord;
}


function buildSessionStatistics(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const sessionIndex = headerRow.indexOf("節次");
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const maxSessionCount = parameterSheet.getRange("B5").getValue();

  const sessionRecords = [];
  for(let sessionNumber = 0; sessionNumber < maxSessionCount + 2; sessionNumber++){
    sessionRecords.push(createEmptySessionRecord());
  }

  for(const studentRow of candidateRows){
    sessionRecords[studentRow[sessionIndex]].students.push(studentRow);
  }

  return sessionRecords;
}
