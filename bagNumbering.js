function bag_numbering() {
  sort_by_session_classroom();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const class_column = headers.indexOf("班級");
  const subject_column = headers.indexOf("科目名稱");
  const session_column = headers.indexOf("節次");
  const classroom_column = headers.indexOf("試場");
  const small_bag_column = headers.indexOf("小袋序號");
  const big_bag_column = headers.indexOf("大袋序號");

  const small_bag_serial = {};
  const big_bag_serial = {};
  let pre_small_number = 0;
  let pre_big_number = 0;

  // 遍歷過所有補考學生-科目，將大小袋編號
  data.forEach(
    function(row){
      let big_key = row[session_column] + "-" + row[classroom_column];
      let small_key = row[session_column] + "-" + row[classroom_column] + "_" + row[class_column] + "=" + row[subject_column];

      if(!Object.keys(big_bag_serial).includes(big_key)){
        big_bag_serial[big_key] = pre_big_number + 1;
        pre_big_number = pre_big_number + 1;
      }

      if(!Object.keys(small_bag_serial).includes(small_key)){
        small_bag_serial[small_key] = pre_small_number + 1;
        pre_small_number = pre_small_number + 1;
      }
    }
  );

  // 將編號寫入記憶體中的陣列
  data.forEach(
    function(row){
      let big_key = row[session_column] + "-" + row[classroom_column];
      let small_key = row[session_column] + "-" + row[classroom_column] + "_" + row[class_column] + "=" + row[subject_column];
      row[big_bag_column] = big_bag_serial[big_key];
      row[small_bag_column] = small_bag_serial[small_key];
    }
  );

  // 將資料寫回試算表
  set_range_values(filtered_sheet.getRange(2, 1, data.length, data[0].length), data);
}
