function generate_big_bag_data(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const big_bag_sheet = ss.getSheetByName("大袋封面套印用資料");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();

  const school_year = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const make_up_date = parametersSheet.getRange("B13").getValue();

  const big_bag_serial_column = headers.indexOf("大袋序號");
  const small_bag_serial_column = headers.indexOf("小袋序號");
  const session_column = headers.indexOf("節次");
  const time_column = headers.indexOf("時間");
  const classroom_column = headers.indexOf("試場");
  const teacher_column = headers.indexOf("監考教師");
  const big_bag_population_column = headers.indexOf("大袋人數");

  big_bag_sheet.clear();

  // 刪除多餘的欄和列，並設置標題列
  if (big_bag_sheet.getMaxRows()>5){
    big_bag_sheet.deleteRows(2,big_bag_sheet.getMaxRows()-5)
  }

  let big_bags = [["學年度", "學期", "大袋序號", "節次", "試場", "補考日期", "時間", "試卷袋序號", "監考教師", "各試場人數"],];
  let already_arranged = [];

  let container = {};
  data.forEach(
    function (row){
      if (!Object.keys(container).includes("大袋" + row[big_bag_serial_column])){
        container["大袋" + row[big_bag_serial_column]] = [row[small_bag_serial_column]]
      } else {
        container["大袋" + row[big_bag_serial_column]].push(row[small_bag_serial_column]);
      }
    }
  );

  data.forEach(
    function(row){
      if (!already_arranged.includes(row[big_bag_serial_column])){
        let tmp = [
          school_year,
          semester,
          row[big_bag_serial_column],  // 大袋序號
          row[session_column],  // 節次
          row[classroom_column], // 試場
          make_up_date,  // 補考日期
          row[time_column],  // 時間
          Math.min(...container["大袋" + row[big_bag_serial_column]]) + "-" + Math.max(...container["大袋" + row[big_bag_serial_column]]),
          row[teacher_column],
          row[big_bag_population_column],
        ];

        big_bags.push(tmp);
        already_arranged.push(row[big_bag_serial_column]);
      }
    }
  );

  set_range_values(big_bag_sheet.getRange(1, 1, big_bags.length, big_bags[0].length), big_bags);
}
