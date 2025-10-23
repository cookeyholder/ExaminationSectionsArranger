function generate_small_bag_data(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const small_bag_sheet = ss.getSheetByName("小袋封面套印用資料");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();

  const school_year = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  
  const small_bag_serial_column = headers.indexOf("小袋序號");
  const session_column = headers.indexOf("節次");
  const time_column = headers.indexOf("時間");
  const classroom_column = headers.indexOf("試場");
  const class_column = headers.indexOf("班級");
  const subject_name_column = headers.indexOf("科目名稱");
  const teacher_column = headers.indexOf("任課老師");
  const small_bag_population_column = headers.indexOf("小袋人數");
  const bycomputer_column = headers.indexOf("電腦");
  const byhand_column = headers.indexOf("人工");

  small_bag_sheet.clear();

  // 刪除多餘的欄和列，並設置標題列
  if (small_bag_sheet.getMaxRows()>5){
    small_bag_sheet.deleteRows(2,small_bag_sheet.getMaxRows()-5)
  }

  let small_bags = [["學年度", "學期", "小袋序號", "節次", "時間", "試場", "班級", "科目名稱", "任課老師", "小袋人數", "電腦", "人工"],];
  let already_arranged = [];

  data.forEach(
    function(row){
      if (!already_arranged.includes(row[small_bag_serial_column])){
        let tmp = [
          school_year,
          semester,
          row[small_bag_serial_column],
          row[session_column],
          row[time_column],
          row[classroom_column],
          row[class_column],
          row[subject_name_column],
          row[teacher_column],
          row[small_bag_population_column],
          row[bycomputer_column],
          row[byhand_column]
        ];

        small_bags.push(tmp);
        already_arranged.push(row[small_bag_serial_column]);
      }
    }
  );

  set_range_values(small_bag_sheet.getRange(1, 1, small_bags.length, small_bags[0].length), small_bags);
}