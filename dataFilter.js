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
  )

  filteredSheet.getRange(2, 1, nameList.length, nameList[0].length)
    .setNumberFormat('@STRING@')  // 改成純文字格式，以免 0 開頭的學號被去掉前面的 0，造成位數錯誤
    .setValues(nameList);
}
