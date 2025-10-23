function generate_bulletin() {
  sort_by_classname()

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parameters_sheet = ss.getSheetByName("參數區");
  const bulletin_sheet = ss.getSheetByName("公告版補考場次");
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();

  const class_column = headers.indexOf("班級");
  const std_number_column = headers.indexOf("學號");
  const name_column = headers.indexOf("姓名");
  const subject_column = headers.indexOf("科目名稱");
  const session_column = headers.indexOf("節次");
  const classroom_column = headers.indexOf("試場");

  // 刪除多餘的欄和列
  bulletin_sheet.clear();
  if (bulletin_sheet.getMaxRows()>5){
    bulletin_sheet.deleteRows(2,bulletin_sheet.getMaxRows() - 5)
  }

  let modified_data = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
  data.forEach(
    function(row){
      let repeat_times = 0;
      let masked_name = "";

      if(row[name_column].length == 2){
        masked_name = row[name_column].toString().slice(0,1) + "〇";
      } else {
        repeat_times = row[name_column].length - 2;
        masked_name = row[name_column].toString().slice(0,1) + "〇".repeat(repeat_times) + row[name_column].toString().slice(-1);
      }

      let tmp = [
        row[class_column],
        row[std_number_column],
        masked_name,
        row[subject_column],
        row[session_column],
        row[classroom_column],
      ]

      modified_data.push(tmp);
    }
  );

  set_range_values(bulletin_sheet.getRange(2, 1, modified_data.length, modified_data[0].length), modified_data);
  sort_by_session_classroom();
  prettier();
}


function prettier(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bulletin_sheet = ss.getSheetByName("公告版補考場次");
  const parameters_sheet = ss.getSheetByName("參數區");

  const school_year = parameters_sheet.getRange("B2").getValue();
  const semester = parameters_sheet.getRange("B3").getValue();

  bulletin_sheet.getRange("A1:F1").mergeAcross();
  bulletin_sheet.getRange("A1").setValue("高雄高工" + school_year + "學年度第" + semester + "學期補考名單");
  bulletin_sheet.getRange("A1").setFontSize(20);
  bulletin_sheet.getRange(1, 1, bulletin_sheet.getMaxRows(), bulletin_sheet.getMaxColumns()).setHorizontalAlignment("center");
  bulletin_sheet.setFrozenRows(2);
  bulletin_sheet.getRange("A2:F").createFilter();
  bulletin_sheet.getRange("A2:F").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}
