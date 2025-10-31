/**
 * 安排共同科目的節次（物理、國文、英文、數學等）
 * 
 * 根據「參數區」工作表中的科目-節次對映規則，
 * 將共同科目分配到指定的節次。
 */
function scheduleCommonSubjectSessions(){
  const sessionRuleRows = PARAMETERS_SHEET.getRange(2, 5, 21, 2).getValues();
  const exam = createExamFromSheet();
  const columns = getColumnIndices();

  // 建立科目到節次的對映
  const preferredSessionBySubject = {};
  sessionRuleRows.forEach(function(ruleRow){
    if (ruleRow[0] && ruleRow[1]){
      preferredSessionBySubject[ruleRow[0]] = ruleRow[1];
    }
  });

  // 重新分配共同科目的節次
  exam.sessions.forEach(function(session){
    session.students.forEach(function(student){
      const subjectName = student[columns.subject];
      const preferredSession = preferredSessionBySubject[subjectName];
      if (preferredSession != null){
        student[columns.session] = preferredSession;
      }
    });
  });

  // 儲存更新後的資料
  saveExamToSheet(exam);
}


/**
 * 安排專業科目的節次
 * 
 * 根據科別年級互斥規則和節次容量限制，
 * 將專業科目分配到適當的節次。
 */
function scheduleSpecializedSubjectSessions(){
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
  const sessionCapacity = 0.9 * PARAMETERS_SHEET.getRange("B9").getValue();

  const departmentGradeSubjectCounts = Object.entries(
    fetchDepartmentGradeSubjectCounts()
  ).sort(compareCountDescending);

  // 清空所有節次（重新分配）
  exam.sessions.forEach(function(session){ 
    session.clear(); 
  });
  
  // 收集所有未分配節次的學生
  const allStudents = [];
  exam.sessions[0].students.forEach(function(student){
    if (student[columns.session] === 0) {
      allStudents.push(student);
    }
  });

  // 為每個節次分配學生
  for (let sessionNumber = 1; sessionNumber <= maxSessionCount; sessionNumber++) {
    const session = exam.sessions[sessionNumber];
    
    for (let countIndex = 0; countIndex < departmentGradeSubjectCounts.length; countIndex++) {
      const [deptGradeSubjectKey, studentCount] = departmentGradeSubjectCounts[countIndex];
      const deptGradeKey = deptGradeSubjectKey.substring(0, deptGradeSubjectKey.indexOf("_"));

      // 檢查該科別年級是否已排入此節次（互斥規則）
      const deptGradeStats = session.departmentGradeStatistics;
      if (Object.keys(deptGradeStats).includes(deptGradeKey)) {
        continue;
      }

      // 檢查容量
      if (studentCount + session.population > sessionCapacity) {
        continue;
      }

      // 分配學生到此節次
      allStudents.forEach(function(student){
        const studentKey = student[columns.department] + student[columns.grade] + "_" + student[columns.subject];
        if (studentKey === deptGradeSubjectKey && student[columns.session] === 0) {
          student[columns.session] = sessionNumber;
          session.addStudent(student);
        }
      });
    }

    if (session.population >= sessionCapacity) {
      Logger.log("第" + sessionNumber + "節已達人數上限：" + session.population);
    }
  }

  // 檢查是否所有學生都已分配
  const unscheduledCount = allStudents.filter(function(s){ 
    return s[columns.session] === 0; 
  }).length;
  
  if (unscheduledCount > 0) {
    Logger.log("無法將所有人排入 " + maxSessionCount + " 節，請檢查是否有某科年級須補考過多科目！");
    SpreadsheetApp.getUi().alert("無法將所有人排入 " + maxSessionCount + " 節，請檢查是否有某科年級須補考過多科目！");
  }

  saveExamToSheet(exam);
}


/**
 * 安排試場
 * 
 * 根據班級科目統計和試場容量限制，
 * 將學生分配到試場。
 */
function assignExamRooms(){
  const exam = createExamFromSheet();
  const columns = getColumnIndices();
  const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();
  const maxRoomCount = PARAMETERS_SHEET.getRange("B6").getValue();
  const maxStudentsPerRoom = PARAMETERS_SHEET.getRange("B7").getValue();
  const maxSubjectsPerRoom = PARAMETERS_SHEET.getRange("B8").getValue();

  for (let sessionNumber = 1; sessionNumber <= maxSessionCount; sessionNumber++) {
    const session = exam.sessions[sessionNumber];
    
    // 清空所有試場
    session.classrooms.forEach(function(classroom){ 
      classroom.clear(); 
    });

    const deptClassSubjectCounts = Object.entries(
      session.departmentClassSubjectStatistics
    ).sort(compareCountDescending);

    for (let roomNumber = 1; roomNumber <= maxRoomCount; roomNumber++) {
      const classroom = session.classrooms[roomNumber];
      let scheduledSubjects = [];

      for (let countIndex = 0; countIndex < deptClassSubjectCounts.length; countIndex++) {
        const [classSubjectKey, count] = deptClassSubjectCounts[countIndex];

        // 檢查是否已排入
        if (scheduledSubjects.includes(classSubjectKey)) {
          continue;
        }

        // 檢查容量
        if (count + classroom.population > maxStudentsPerRoom) {
          continue;
        }

        // 檢查科目數限制
        const subjectCount = Object.keys(classroom.classSubjectStatistics).length;
        if (subjectCount + 1 > maxSubjectsPerRoom) {
          continue;
        }

        // 分配學生到此試場
        session.students.forEach(function(student){
          const studentKey = student[columns.class] + student[columns.subject];
          if (studentKey === classSubjectKey && student[columns.room] === 0) {
            student[columns.room] = roomNumber;
            classroom.addStudent(student);
          }
        });

        scheduledSubjects = Object.keys(classroom.classSubjectStatistics);
      }
    }
  }

  // 檢查是否所有學生都已分配試場
  let allScheduled = true;
  exam.sessions.forEach(function(session){
    session.students.forEach(function(student){
      if (student[columns.room] === 0) {
        allScheduled = false;
      }
    });
  });

  if (!allScheduled) {
    Logger.log("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
    SpreadsheetApp.getUi().alert("現有試場數無法容納所有補考學生，請增加試場數或調整每間試場人數上限！");
  }

  saveExamToSheet(exam);

  // 檢查是否有學生被安排在第 9 節
  if (exam.sessions[9] && exam.sessions[9].population > 0) {
    SpreadsheetApp.getUi().alert("部分考生被安排在第9節補考，請注意是否需要調整到中午應試！");
  }

  FILTERED_RESULT_SHEET.getRange("I:J").setNumberFormat('#,##0');
  sortFilteredStudentsBySessionRoom();
}


function allocateBagIdentifiers() {
  sortFilteredStudentsBySessionRoom();

  const [headerRow, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();
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

  writeRangeValuesSafely(FILTERED_RESULT_SHEET.getRange(2, 1, candidateRows.length, candidateRows[0].length), candidateRows);
}


function populateSessionTimes() {
  const [headerRow, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  const sessionIndex = headerRow.indexOf("節次");
  const timeIndex = headerRow.indexOf("時間");

  const [timeHeaders, ...sessionTimeRows] = SESSION_TIME_REFERENCE_SHEET.getDataRange().getValues();

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
    writeRangeValuesSafely(FILTERED_RESULT_SHEET.getRange(2, 1, updatedRows.length, updatedRows[0].length), updatedRows);
  } else {
    Logger.log("寫入節次時間失敗！");
    SpreadsheetApp.getUi().alert("寫入節次時間失敗！");
  }
}


function updateBagAndClassPopulations() {
  const [headerRow, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();
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

  writeRangeValuesSafely(FILTERED_RESULT_SHEET.getRange(2, 1, candidateRows.length, candidateRows[0].length), candidateRows);
}


function sortFilteredStudentsBySubject(){
  const filteredRange = FILTERED_RESULT_SHEET.getDataRange();

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
  const filteredRange = FILTERED_RESULT_SHEET.getDataRange();

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
  const filteredRange = FILTERED_RESULT_SHEET.getDataRange();

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
  const [, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();

  return summarizeDepartmentGradeCounts(candidateRows);
}


function fetchDepartmentGradeSubjectCounts(){
  const [, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();

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
  const maxRoomCount = PARAMETERS_SHEET.getRange("B6").getValue();
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
  const [headerRow, ...candidateRows] = FILTERED_RESULT_SHEET.getDataRange().getValues();
  const sessionIndex = headerRow.indexOf("節次");
  const maxSessionCount = PARAMETERS_SHEET.getRange("B5").getValue();

  const sessionRecords = [];
  for(let sessionNumber = 0; sessionNumber < maxSessionCount + 2; sessionNumber++){
    sessionRecords.push(createEmptySessionRecord());
  }

  for(const studentRow of candidateRows){
    sessionRecords[studentRow[sessionIndex]].students.push(studentRow);
  }

  return sessionRecords;
}
