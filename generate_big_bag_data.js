function generateBigBagData(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const bigBagSheet = ss.getSheetByName("大袋封面套印用資料");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();

  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const makeUpDate = parametersSheet.getRange("B13").getValue();

  const bigBagSerialColumn = headers.indexOf("大袋序號");
  const smallBagSerialColumn = headers.indexOf("小袋序號");
  const sessionColumn = headers.indexOf("節次");
  const timeColumn = headers.indexOf("時間");
  const classroomColumn = headers.indexOf("試場");
  const teacherColumn = headers.indexOf("監考教師");
  const bigBagPopulationColumn = headers.indexOf("大袋人數");

  bigBagSheet.clear();

  // 刪除多餘的欄和列，並設置標題列
  if (bigBagSheet.getMaxRows()>5){
    bigBagSheet.deleteRows(2,bigBagSheet.getMaxRows()-5)
  }

  let bigBags = [["學年度", "學期", "大袋序號", "節次", "試場", "補考日期", "時間", "試卷袋序號", "監考教師", "各試場人數"],];
  let alreadyArranged = [];

  let container = {};
  data.forEach(
    function (row){
      if (!Object.keys(container).includes("大袋" + row[bigBagSerialColumn])){
        container["大袋" + row[bigBagSerialColumn]] = [row[smallBagSerialColumn]]
      } else {
        container["大袋" + row[bigBagSerialColumn]].push(row[smallBagSerialColumn]);
      }
    }
  );

  data.forEach(
    function(row){
      if (!alreadyArranged.includes(row[bigBagSerialColumn])){
        let tmp = [
          schoolYear,
          semester,
          row[bigBagSerialColumn],  // 大袋序號
          row[sessionColumn],  // 節次
          row[classroomColumn], // 試場
          makeUpDate,  // 補考日期
          row[timeColumn],  // 時間
          Math.min(...container["大袋" + row[bigBagSerialColumn]]) + "-" + Math.max(...container["大袋" + row[bigBagSerialColumn]]),
          row[teacherColumn],
          row[bigBagPopulationColumn],
        ];

        bigBags.push(tmp);
        alreadyArranged.push(row[bigBagSerialColumn]);
      }
    }
  );

  setRangeValues(bigBagSheet.getRange(1, 1, bigBags.length, bigBags[0].length), bigBags);
}
