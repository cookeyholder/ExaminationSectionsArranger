function calculateClassroomPopulation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const sessionColumn = headers.indexOf("節次");
  const classColumn = headers.indexOf("班級");
  const bigBagSerialColumn = headers.indexOf("大袋序號");
  const smallBagSerialColumn = headers.indexOf("小袋序號");
  const bigBagPopulationColumn = headers.indexOf("大袋人數");
  const smallBagPopulationColumn = headers.indexOf("小袋人數");
  const classPopulationColumn = headers.indexOf("班級人數");

  let bigBagPopulation = {};
  let smallBagPopulation = {};
  let classPopulation = {};
  // 遍歷所有補考學生-科目，數一數各大袋、小袋裡的人數
  data.forEach(
    function(row){
      let bigBagKey = "大袋" + row[bigBagSerialColumn];
      if(Object.keys(bigBagPopulation).includes(bigBagKey)){
        bigBagPopulation[bigBagKey] += 1;
      } else {
        bigBagPopulation[bigBagKey] = 1;
      }

      let smallBagKey = "小袋" + row[smallBagSerialColumn];
      if(Object.keys(smallBagPopulation).includes(smallBagKey)){
        smallBagPopulation[smallBagKey] += 1;
      } else {
        smallBagPopulation[smallBagKey] = 1;
      }

      let classPopulationKey = "班級" + row[smallBagSerialColumn] + row[classColumn];
      if(Object.keys(classPopulation).includes(classPopulationKey)){
        classPopulation[classPopulationKey] += 1;
      } else {
        classPopulation[classPopulationKey] = 1;
      }
    }
  );

  // 將編號寫入記憶體中的陣列
  data.forEach(
    function(row){
      let bigBagKey = "大袋" + row[bigBagSerialColumn];
      let smallBagKey = "小袋" + row[smallBagSerialColumn];
      let classPopulationKey = "班級" + row[smallBagSerialColumn] + row[classColumn];
      row[bigBagPopulationColumn] = bigBagPopulation[bigBagKey];
      row[smallBagPopulationColumn] = smallBagPopulation[smallBagKey];
      row[classPopulationColumn] = classPopulation[classPopulationKey];
    }
  );

  // 將資料寫回試算表
  setRangeValues(filteredSheet.getRange(2, 1, data.length, data[0].length), data);
}
