function descendingSorting(a, b) {
  if (a[1] === b[1]) {
    return 0;
  } else {
    return (a[1] < b[1]) ? 1 : -1;
  }
}


function getDepartmentGradeStatisticsOfArray(data){
  const departmentColumn = 0;
  const gradeColumn = 1;
  const statistics = {};
  for (const row of data){
    const key = row[departmentColumn] + row[gradeColumn];

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
  
  const statistics = {};
  for (const row of data){
    const key = row[departmentColumn] + row[gradeColumn] + "_" + row[subjectNameColumn];

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
      const statistics = {};
      this.students.forEach(
        function(row){
          const key = row[3] + "_" + row[7];  // 班級 + _ + 科目
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
      const statistics = {};
      this.students.forEach(
        function(row){
          const key = row[0]+row[1];  // 科別 + 年級
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
        function(row){
          const key = row[3] + row[7];  // 班級 + 科目
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
    session.classrooms.push(createClassroom());
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
    sessions.push(createSession());
  }

  for(const row of data){
    sessions[row[sessionColumn]].students.push(row);
  }

  return sessions;
}


function arrangeCommonsSession(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const [paramHeaders,...commonData] = parametersSheet.getRange(2, 5, 21, 2).getValues();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const subjectColumn = headers.indexOf("科目名稱");
  const sessionColumn = headers.indexOf("節次");

  const commonSessions = {};
  commonData.forEach(
    function(row){
      commonSessions[row[0]] = row[1];
    }
  );
  
  const modifiedData = data.map(
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
            const key = row[0] + row[1] + "_" + row[7];
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
              const key = row[classColumn] + row[subjectColumn];
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
      for(let sessionIdx=1; sessionIdx < MAX_SESSION_NUMBER + 2; sessionIdx++){
        for(let classroomIdx=1; classroomIdx < sessions[sessionIdx].classrooms.length; classroomIdx++){
          mergedSessionClassrooms = mergedSessionClassrooms.concat(sessions[sessionIdx].classrooms[classroomIdx].students);
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


function bagNumbering() {
  sortBySessionClassroom();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const classColumn = headers.indexOf("班級");
  const subjectColumn = headers.indexOf("科目名稱");
  const sessionColumn = headers.indexOf("節次");
  const classroomColumn = headers.indexOf("試場");
  const smallBagColumn = headers.indexOf("小袋序號");
  const bigBagColumn = headers.indexOf("大袋序號");

  const smallBagSerial = {};
  const bigBagSerial = {};
  let previousSmallNumber = 0;
  let previousBigNumber = 0;

  // 遍歷過所有補考學生-科目，將大小袋編號
  data.forEach(
    function(row){
      const bigKey = row[sessionColumn] + "-" + row[classroomColumn];
      const smallKey = row[sessionColumn] + "-" + row[classroomColumn] + "_" + row[classColumn] + "=" + row[subjectColumn];

      if(!Object.keys(bigBagSerial).includes(bigKey)){
        bigBagSerial[bigKey] = previousBigNumber + 1;
        previousBigNumber = previousBigNumber + 1;
      }

      if(!Object.keys(smallBagSerial).includes(smallKey)){
        smallBagSerial[smallKey] = previousSmallNumber + 1;
        previousSmallNumber = previousSmallNumber + 1;
      }
    }
  );

  // 將編號寫入記憶體中的陣列
  data.forEach(
    function(row){
      const bigKey = row[sessionColumn] + "-" + row[classroomColumn];
      const smallKey = row[sessionColumn] + "-" + row[classroomColumn] + "_" + row[classColumn] + "=" + row[subjectColumn];
      row[bigBagColumn] = bigBagSerial[bigKey];
      row[smallBagColumn] = smallBagSerial[smallKey];
    }
  );

  // 將資料寫回試算表
  setRangeValues(filteredSheet.getRange(2, 1, data.length, data[0].length), data);
}


function calculateClassroomPopulation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const classColumn = headers.indexOf("班級");
  const bigBagSerialColumn = headers.indexOf("大袋序號");
  const smallBagSerialColumn = headers.indexOf("小袋序號");
  const bigBagPopulationColumn = headers.indexOf("大袋人數");
  const smallBagPopulationColumn = headers.indexOf("小袋人數");
  const classPopulationColumn = headers.indexOf("班級人數");

  const bigBagPopulation = {};
  const smallBagPopulation = {};
  const classPopulation = {};
  // 遍歷所有補考學生-科目，數一數各大袋、小袋裡的人數
  data.forEach(
    function(row){
      const bigBagKey = "大袋" + row[bigBagSerialColumn];
      if(Object.keys(bigBagPopulation).includes(bigBagKey)){
        bigBagPopulation[bigBagKey] += 1;
      } else {
        bigBagPopulation[bigBagKey] = 1;
      }

      const smallBagKey = "小袋" + row[smallBagSerialColumn];
      if(Object.keys(smallBagPopulation).includes(smallBagKey)){
        smallBagPopulation[smallBagKey] += 1;
      } else {
        smallBagPopulation[smallBagKey] = 1;
      }

      const classPopulationKey = "班級" + row[smallBagSerialColumn] + row[classColumn];
      if(Object.keys(classPopulation).includes(classPopulationKey)){
        classPopulation[classPopulationKey] += 1;
      } else {
        classPopulation[classPopulationKey] = 1;
      }
    }
  );

  // 將編號寫入記憶體中的陣列
  data.forEach(
    function(row){
      const bigBagKey = "大袋" + row[bigBagSerialColumn];
      const smallBagKey = "小袋" + row[smallBagSerialColumn];
      const classPopulationKey = "班級" + row[smallBagSerialColumn] + row[classColumn];
      row[bigBagPopulationColumn] = bigBagPopulation[bigBagKey];
      row[smallBagPopulationColumn] = smallBagPopulation[smallBagKey];
      row[classPopulationColumn] = classPopulation[classPopulationKey];
    }
  );

  // 將資料寫回試算表
  setRangeValues(filteredSheet.getRange(2, 1, data.length, data[0].length), data);
}


function setSessionTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const sessionColumn = headers.indexOf("節次");
  const timeColumn = headers.indexOf("時間");

  const timeTableSheet = ss.getSheetByName("節次時間表");
  const [timeHeaders, ...timeData] = timeTableSheet.getDataRange().getValues();

  const timeTable = {};
  timeData.forEach(
    function(timeRow){
      timeTable[timeRow[0]] = timeRow[1];
    }
  );

  const modifiedData = data.map(
    function(row){
      row[timeColumn] = timeTable[row[sessionColumn]];
      return row;
    }
  );

  if(modifiedData.length == data.length){
    setRangeValues(filteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("寫入節次時間失敗！");
    SpreadsheetApp.getUi().alert("寫入節次時間失敗！");
  }
}


function sortBySubject(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filteredSheet.getDataRange();

  filtered.offset(1,0,filtered.getNumRows()-1).sort(
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


function sortByClassname(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filteredSheet.getDataRange();

  filtered.offset(1,0,filtered.getNumRows()-1).sort(
    [
      {column: 2, ascending: true}, 
      {column: 3, ascending: true}, 
      {column: 5, ascending: true}, 
      {column: 9, ascending: true}, 
      {column: 8, ascending: true}
    ]
  );
}


function sortBySessionClassroom(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filteredSheet.getDataRange();

  filtered.offset(1,0,filtered.getNumRows()-1).sort(
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
