function createExamBulletinSheet() {
  sortFilteredStudentsByClassSeat();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const bulletinSheet = spreadsheet.getSheetByName("公告版補考場次");
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();

  const classIndex = headerRow.indexOf("班級");
  const studentNumberIndex = headerRow.indexOf("學號");
  const studentNameIndex = headerRow.indexOf("姓名");
  const subjectIndex = headerRow.indexOf("科目名稱");
  const sessionIndex = headerRow.indexOf("節次");
  const roomIndex = headerRow.indexOf("試場");

  bulletinSheet.clear();
  if (bulletinSheet.getMaxRows()>5){
    bulletinSheet.deleteRows(2,bulletinSheet.getMaxRows() - 5);
  }

  const bulletinRows = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
  candidateRows.forEach(
    function(examineeRow){
      let maskedName = "";
      if(examineeRow[studentNameIndex].length === 2){
        maskedName = examineeRow[studentNameIndex].toString().slice(0,1) + "〇";
      } else {
        const middleMaskLength = examineeRow[studentNameIndex].length - 2;
        maskedName = examineeRow[studentNameIndex].toString().slice(0,1) + "〇".repeat(middleMaskLength) + examineeRow[studentNameIndex].toString().slice(-1);
      }

      const bulletinRow = [
        examineeRow[classIndex],
        examineeRow[studentNumberIndex],
        maskedName,
        examineeRow[subjectIndex],
        examineeRow[sessionIndex],
        examineeRow[roomIndex],
      ];
      bulletinRows.push(bulletinRow);
    }
  );

  writeRangeValuesSafely(bulletinSheet.getRange(2, 1, bulletinRows.length, bulletinRows[0].length), bulletinRows);
  sortFilteredStudentsBySessionRoom();
  formatBulletinSheet();
}


function formatBulletinSheet(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const bulletinSheet = spreadsheet.getSheetByName("公告版補考場次");
  const parameterSheet = spreadsheet.getSheetByName("參數區");

  const schoolYearValue = parameterSheet.getRange("B2").getValue();
  const semesterValue = parameterSheet.getRange("B3").getValue();

  bulletinSheet.getRange("A1:F1").mergeAcross();
  bulletinSheet.getRange("A1").setValue("高雄高工" + schoolYearValue + "學年度第" + semesterValue + "學期補考名單");
  bulletinSheet.getRange("A1").setFontSize(20);
  bulletinSheet.getRange(1, 1, bulletinSheet.getMaxRows(), bulletinSheet.getMaxColumns()).setHorizontalAlignment("center");
  bulletinSheet.setFrozenRows(2);
  bulletinSheet.getRange("A2:F").createFilter();
  bulletinSheet.getRange("A2:F").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}


function createProctorRecordSheet() {
  sortFilteredStudentsBySessionRoom();

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const recordSheet = spreadsheet.getSheetByName("試場紀錄表(A表)");

  const schoolYearValue = parameterSheet.getRange("B2").getValue();
  const semesterValue = parameterSheet.getRange("B3").getValue();
  const sessionIndex = headerRow.indexOf("節次");
  const roomIndex = headerRow.indexOf("試場");
  const timeIndex = headerRow.indexOf("時間");
  const classIndex = headerRow.indexOf("班級");
  const studentNumberIndex = headerRow.indexOf("學號");
  const studentNameIndex = headerRow.indexOf("姓名");
  const subjectIndex = headerRow.indexOf("科目名稱");
  const classPopulationIndex = headerRow.indexOf("班級人數");

  recordSheet.clear();
  if (recordSheet.getMaxRows()>5){
    recordSheet.deleteRows(2,recordSheet.getMaxRows() - 5);
  }

  const recordRows = [
    ["A表："+schoolYearValue+"學年度第"+semesterValue+"學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　", "", "", "", "", "", "", "", "", "", "", "",],
    ["節次", "試場",	"時間",	"班級",	"學號",	"姓名",	"科目名稱",	"班級人數",	"考生到考簽名", "違規記錄(打V)", "", "其他違規\n請簡述"],
    ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""]
  ];

  candidateRows.forEach(
    function(examineeRow){
      recordRows.push([
        examineeRow[sessionIndex],
        examineeRow[roomIndex],
        examineeRow[timeIndex],
        examineeRow[classIndex],
        examineeRow[studentNumberIndex],
        examineeRow[studentNameIndex],
        examineeRow[subjectIndex],
        examineeRow[classPopulationIndex],
        "",
        "",
        "",
        ""
      ]);
    }
  );

  writeRangeValuesSafely(recordSheet.getRange(1, 1, recordRows.length, recordRows[0].length), recordRows);

  recordSheet.getRange("A1:L1").mergeAcross().setVerticalAlignment("bottom").setFontSize(14).setFontWeight("bold");
  recordSheet.getRange("J2:K2").mergeAcross();
  recordSheet.getRange("A2:A3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("B2:B3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("C2:C3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("D2:D3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("E2:E3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("F2:F3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("G2:G3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("H2:H3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("I2:I3").mergeVertically().setVerticalAlignment("middle");
  recordSheet.getRange("L2:L3").mergeVertically().setVerticalAlignment("middle");

  recordSheet.getRange(2, 1, recordRows.length + 2, recordRows[0].length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}


function composeSmallBagDataset(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const smallBagSheet = spreadsheet.getSheetByName("小袋封面套印用資料");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();

  const schoolYearValue = parameterSheet.getRange("B2").getValue();
  const semesterValue = parameterSheet.getRange("B3").getValue();
  
  const smallBagIndex = headerRow.indexOf("小袋序號");
  const sessionIndex = headerRow.indexOf("節次");
  const timeIndex = headerRow.indexOf("時間");
  const roomIndex = headerRow.indexOf("試場");
  const classIndex = headerRow.indexOf("班級");
  const subjectIndex = headerRow.indexOf("科目名稱");
  const teacherIndex = headerRow.indexOf("任課老師");
  const smallBagPopulationIndex = headerRow.indexOf("小袋人數");
  const computerIndex = headerRow.indexOf("電腦");
  const manualIndex = headerRow.indexOf("人工");

  smallBagSheet.clear();
  if (smallBagSheet.getMaxRows()>5){
    smallBagSheet.deleteRows(2,smallBagSheet.getMaxRows()-5);
  }

  const smallBagRows = [["學年度", "學期", "小袋序號", "節次", "時間", "試場", "班級", "科目名稱", "任課老師", "小袋人數", "電腦", "人工"],];
  const processedSmallBags = [];

  candidateRows.forEach(
    function(examineeRow){
      if (!processedSmallBags.includes(examineeRow[smallBagIndex])){
        const datasetRow = [
          schoolYearValue,
          semesterValue,
          examineeRow[smallBagIndex],
          examineeRow[sessionIndex],
          examineeRow[timeIndex],
          examineeRow[roomIndex],
          examineeRow[classIndex],
          examineeRow[subjectIndex],
          examineeRow[teacherIndex],
          examineeRow[smallBagPopulationIndex],
          examineeRow[computerIndex],
          examineeRow[manualIndex]
        ];

        smallBagRows.push(datasetRow);
        processedSmallBags.push(examineeRow[smallBagIndex]);
      }
    }
  );

  writeRangeValuesSafely(smallBagSheet.getRange(1, 1, smallBagRows.length, smallBagRows[0].length), smallBagRows);
}


function composeBigBagDataset(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const bigBagSheet = spreadsheet.getSheetByName("大袋封面套印用資料");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();

  const schoolYearValue = parameterSheet.getRange("B2").getValue();
  const semesterValue = parameterSheet.getRange("B3").getValue();
  const makeUpDateValue = parameterSheet.getRange("B13").getValue();

  const bigBagIndex = headerRow.indexOf("大袋序號");
  const smallBagIndex = headerRow.indexOf("小袋序號");
  const sessionIndex = headerRow.indexOf("節次");
  const timeIndex = headerRow.indexOf("時間");
  const roomIndex = headerRow.indexOf("試場");
  const invigilatorIndex = headerRow.indexOf("監考教師");
  const bigBagPopulationIndex = headerRow.indexOf("大袋人數");

  bigBagSheet.clear();
  if (bigBagSheet.getMaxRows()>5){
    bigBagSheet.deleteRows(2,bigBagSheet.getMaxRows()-5);
  }

  const bigBagRows = [["學年度", "學期", "大袋序號", "節次", "試場", "補考日期", "時間", "試卷袋序號", "監考教師", "各試場人數"],];
  const processedBigBags = [];

  const smallBagRangeByBigBag = {};
  candidateRows.forEach(
    function(examineeRow){
      const bigBagKey = "大袋" + examineeRow[bigBagIndex];
      if (!Object.keys(smallBagRangeByBigBag).includes(bigBagKey)){
        smallBagRangeByBigBag[bigBagKey] = [examineeRow[smallBagIndex]];
      } else {
        smallBagRangeByBigBag[bigBagKey].push(examineeRow[smallBagIndex]);
      }
    }
  );

  candidateRows.forEach(
    function(examineeRow){
      if (!processedBigBags.includes(examineeRow[bigBagIndex])){
        const bagRange = smallBagRangeByBigBag["大袋" + examineeRow[bigBagIndex]];
        const datasetRow = [
          schoolYearValue,
          semesterValue,
          examineeRow[bigBagIndex],
          examineeRow[sessionIndex],
          examineeRow[roomIndex],
          makeUpDateValue,
          examineeRow[timeIndex],
          Math.min(...bagRange) + "-" + Math.max(...bagRange),
          examineeRow[invigilatorIndex],
          examineeRow[bigBagPopulationIndex],
        ];

        bigBagRows.push(datasetRow);
        processedBigBags.push(examineeRow[bigBagIndex]);
      }
    }
  );

  writeRangeValuesSafely(bigBagSheet.getRange(1, 1, bigBagRows.length, bigBagRows[0].length), bigBagRows);
}
