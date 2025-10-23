function generate_record_sheet() {
  sort_by_session_classroom();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const record_sheet = ss.getSheetByName("試場紀錄表(A表)");

  const school_year = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const session_column = headers.indexOf("節次");
  const classroom_column = headers.indexOf("試場");
  const time_column = headers.indexOf("時間");
  const class_column = headers.indexOf("班級");
  const std_number_column = headers.indexOf("學號");
  const name_column = headers.indexOf("姓名");
  const subject_column = headers.indexOf("科目名稱");
  const class_population_column = headers.indexOf("班級人數");

  // 刪除多餘的欄和列
  record_sheet.clear();
  if (record_sheet.getMaxRows()>5){
    record_sheet.deleteRows(2,record_sheet.getMaxRows() - 5)
  }

  let modified_data = [
      ["A表："+school_year+"學年度第"+semester+"學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　", "", "", "", "", "", "", "", "", "", "", "",],
      ["節次", "試場",	"時間",	"班級",	"學號",	"姓名",	"科目名稱",	"班級人數",	"考生到考簽名", "違規記錄(打V)", "", "其他違規\n請簡述"],
      ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""]
    ];

  data.forEach(
    function (row){
      modified_data.push([
        row[session_column],
        row[classroom_column],
        row[time_column],
        row[class_column],
        row[std_number_column],
        row[name_column],
        row[subject_column],
        row[class_population_column],
        "",
        "",
        "",
        ""
      ]);
    }
  );

  set_range_values(record_sheet.getRange(1, 1, modified_data.length, modified_data[0].length), modified_data);

  // 設定格式美化表格
  record_sheet.getRange("A1:L1").mergeAcross().setVerticalAlignment("bottom").setFontSize(14).setFontWeight("bold");
  record_sheet.getRange("J2:K2").mergeAcross();
  record_sheet.getRange("A2:A3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("B2:B3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("C2:C3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("D2:D3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("E2:E3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("F2:F3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("G2:G3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("H2:H3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("I2:I3").mergeVertically().setVerticalAlignment("middle");
  record_sheet.getRange("L2:L3").mergeVertically().setVerticalAlignment("middle");

  record_sheet.getRange(2, 1, modified_data.length + 2, modified_data[0].length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}
