function generateBulletin() {
  sortByClassname()

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const bulletinSheet = ss.getSheetByName("公告版補考場次");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();

  const classColumn = headers.indexOf("班級");
  const stdNumberColumn = headers.indexOf("學號");
  const nameColumn = headers.indexOf("姓名");
  const subjectColumn = headers.indexOf("科目名稱");
  const sessionColumn = headers.indexOf("節次");
  const classroomColumn = headers.indexOf("試場");

  // 刪除多餘的欄和列
  bulletinSheet.clear();
  if (bulletinSheet.getMaxRows()>5){
    bulletinSheet.deleteRows(2,bulletinSheet.getMaxRows() - 5)
  }

  let modifiedData = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
  data.forEach(
    function(row){
      let repeatTimes = 0;
      let maskedName = "";

      if(row[nameColumn].length == 2){
        maskedName = row[nameColumn].toString().slice(0,1) + "〇";
      } else {
        repeatTimes = row[nameColumn].length - 2;
        maskedName = row[nameColumn].toString().slice(0,1) + "〇".repeat(repeatTimes) + row[nameColumn].toString().slice(-1);
      }

      let tmp = [
        row[classColumn],
        row[stdNumberColumn],
        maskedName,
        row[subjectColumn],
        row[sessionColumn],
        row[classroomColumn],
      ]

      modifiedData.push(tmp);
    }
  );

  setRangeValues(bulletinSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  sortBySessionClassroom();
  prettier();
}


function prettier(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bulletinSheet = ss.getSheetByName("公告版補考場次");
  const parametersSheet = ss.getSheetByName("參數區");

  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();

  bulletinSheet.getRange("A1:F1").mergeAcross();
  bulletinSheet.getRange("A1").setValue("高雄高工" + schoolYear + "學年度第" + semester + "學期補考名單");
  bulletinSheet.getRange("A1").setFontSize(20);
  bulletinSheet.getRange(1, 1, bulletinSheet.getMaxRows(), bulletinSheet.getMaxColumns()).setHorizontalAlignment("center");
  bulletinSheet.setFrozenRows(2);
  bulletinSheet.getRange("A2:F").createFilter();
  bulletinSheet.getRange("A2:F").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}
