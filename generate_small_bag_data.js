function generateSmallBagData(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const smallBagSheet = ss.getSheetByName("小袋封面套印用資料");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();

  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  
  const smallBagSerialColumn = headers.indexOf("小袋序號");
  const sessionColumn = headers.indexOf("節次");
  const timeColumn = headers.indexOf("時間");
  const classroomColumn = headers.indexOf("試場");
  const classColumn = headers.indexOf("班級");
  const subjectNameColumn = headers.indexOf("科目名稱");
  const teacherColumn = headers.indexOf("任課老師");
  const smallBagPopulationColumn = headers.indexOf("小袋人數");
  const byComputerColumn = headers.indexOf("電腦");
  const byHandColumn = headers.indexOf("人工");

  smallBagSheet.clear();

  // 刪除多餘的欄和列，並設置標題列
  if (smallBagSheet.getMaxRows()>5){
    smallBagSheet.deleteRows(2,smallBagSheet.getMaxRows()-5)
  }

  let smallBags = [["學年度", "學期", "小袋序號", "節次", "時間", "試場", "班級", "科目名稱", "任課老師", "小袋人數", "電腦", "人工"],];
  let alreadyArranged = [];

  data.forEach(
    function(row){
      if (!alreadyArranged.includes(row[smallBagSerialColumn])){
        let tmp = [
          schoolYear,
          semester,
          row[smallBagSerialColumn],
          row[sessionColumn],
          row[timeColumn],
          row[classroomColumn],
          row[classColumn],
          row[subjectNameColumn],
          row[teacherColumn],
          row[smallBagPopulationColumn],
          row[byComputerColumn],
          row[byHandColumn]
        ];

        smallBags.push(tmp);
        alreadyArranged.push(row[smallBagSerialColumn]);
      }
    }
  );

  setRangeValues(smallBagSheet.getRange(1, 1, smallBags.length, smallBags[0].length), smallBags);
}
