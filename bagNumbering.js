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
      let bigKey = row[sessionColumn] + "-" + row[classroomColumn];
      let smallKey = row[sessionColumn] + "-" + row[classroomColumn] + "_" + row[classColumn] + "=" + row[subjectColumn];

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
      let bigKey = row[sessionColumn] + "-" + row[classroomColumn];
      let smallKey = row[sessionColumn] + "-" + row[classroomColumn] + "_" + row[classColumn] + "=" + row[subjectColumn];
      row[bigBagColumn] = bigBagSerial[bigKey];
      row[smallBagColumn] = smallBagSerial[smallKey];
    }
  );

  // 將資料寫回試算表
  setRangeValues(filteredSheet.getRange(2, 1, data.length, data[0].length), data);
}
