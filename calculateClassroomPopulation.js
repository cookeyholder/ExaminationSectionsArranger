function calculate_classroom_population() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const session_column = headers.indexOf("節次");
  const class_column = headers.indexOf("班級");
  const big_bag_serial_column = headers.indexOf("大袋序號");
  const small_bag_serial_column = headers.indexOf("小袋序號");
  const big_bag_population_column = headers.indexOf("大袋人數");
  const small_bag_population_column = headers.indexOf("小袋人數");
  const class_population_column = headers.indexOf("班級人數");

  let big_bag_population = {};
  let small_bag_population = {};
  let class_population = {};
  // 遍歷所有補考學生-科目，數一數各大袋、小袋裡的人數
  data.forEach(
    function(row){
      let big_bag_key = "大袋" + row[big_bag_serial_column];
      if(Object.keys(big_bag_population).includes(big_bag_key)){
        big_bag_population[big_bag_key] += 1;
      } else {
        big_bag_population[big_bag_key] = 1;
      }

      let small_bag_key = "小袋" + row[small_bag_serial_column];
      if(Object.keys(small_bag_population).includes(small_bag_key)){
        small_bag_population[small_bag_key] += 1;
      } else {
        small_bag_population[small_bag_key] = 1;
      }

      let class_population_key = "班級" + row[small_bag_serial_column] + row[class_column];
      if(Object.keys(class_population).includes(class_population_key)){
        class_population[class_population_key] += 1;
      } else {
        class_population[class_population_key] = 1;
      }
    }
  );

  // 將編號寫入記憶體中的陣列
  data.forEach(
    function(row){
      let big_bag_key = "大袋" + row[big_bag_serial_column];
      let small_bag_key = "小袋" + row[small_bag_serial_column];
      let class_population_key = "班級" + row[small_bag_serial_column] + row[class_column];
      row[big_bag_population_column] = big_bag_population[big_bag_key];
      row[small_bag_population_column] = small_bag_population[small_bag_key];
      row[class_population_column] = class_population[class_population_key];
    }
  );

  // 將資料寫回試算表
  set_range_values(filtered_sheet.getRange(2, 1, data.length, data[0].length), data);
}
