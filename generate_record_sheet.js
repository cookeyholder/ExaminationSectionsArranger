function generateRecordSheet() {
  sortBySessionClassroom();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const recordSheet = ss.getSheetByName("試場紀錄表(A表)");

  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const sessionColumn = headers.indexOf("節次");
  const classroomColumn = headers.indexOf("試場");
  const timeColumn = headers.indexOf("時間");
  const classColumn = headers.indexOf("班級");
  const stdNumberColumn = headers.indexOf("學號");
  const nameColumn = headers.indexOf("姓名");
  const subjectColumn = headers.indexOf("科目名稱");
  const classPopulationColumn = headers.indexOf("班級人數");

  // 刪除多餘的欄和列
  recordSheet.clear();
  if (recordSheet.getMaxRows()>5){
    recordSheet.deleteRows(2,recordSheet.getMaxRows() - 5)
  }

  let modifiedData = [
      ["A表："+schoolYear+"學年度第"+semester+"學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　", "", "", "", "", "", "", "", "", "", "", "",],
      ["節次", "試場",	"時間",	"班級",	"學號",	"姓名",	"科目名稱",	"班級人數",	"考生到考簽名", "違規記錄(打V)", "", "其他違規\n請簡述"],
      ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""]
    ];

  data.forEach(
    function (row){
      modifiedData.push([
        row[sessionColumn],
        row[classroomColumn],
        row[timeColumn],
        row[classColumn],
        row[stdNumberColumn],
        row[nameColumn],
        row[subjectColumn],
        row[classPopulationColumn],
        "",
        "",
        "",
        ""
      ]);
    }
  );

  setRangeValues(recordSheet.getRange(1, 1, modifiedData.length, modifiedData[0].length), modifiedData);

  // 設定格式美化表格
  recordSheet.getRange("A1:L1").mergeAcross().setVerticalAlignment("bottom").setFontSize(14).setFontWeight("bold");
  recordSheet.getRange("J2:K2").mergeAcross();
  recordSheet.getRange("A2:A3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("B2:B3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("C2:C3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("D2:D3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("E2:E3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("F2:F3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("G2:G3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("H2:H3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("I2:I3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("L2:L3").mergeVertically().setVerticalAlignment("middle");

  recordSheet.getRange(2, 1, modifiedData.length + 2, modifiedData[0].length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}
