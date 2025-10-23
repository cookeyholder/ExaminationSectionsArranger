function generateBulletin() {
  sortByClassname();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const bulletinSheet = ss.getSheetByName("公告版補考場次");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();

  const classColumn = headers.indexOf("班級");
  const stdNumberColumn = headers.indexOf("學號");
  const nameColumn = headers.indexOf("姓名");
  const subjectColumn = headers.indexOf("科目名稱");
  const sessionColumn = headers.indexOf("節次");
  const classroomColumn = headers.indexOf("試場");

  // 刪除多餘的欄和列
  bulletinSheet.clear();
  if (bulletinSheet.getMaxRows()>5){
    bulletinSheet.deleteRows(2,bulletinSheet.getMaxRows() - 5);
  }

  const modifiedData = [["班級", "學號", "姓名", "科目", "節次", "試場"]];
  data.forEach(
    function(row){
      let repeatTimes = 0;
      let maskedName = "";

      if(row[nameColumn].length == 2){
        maskedName = row[nameColumn].toString().slice(0,1) + "〇";
      } else {
        repeatTimes = row[nameColumn].length - 2;
        maskedName = row[nameColumn].toString().slice(0,1) + "〇".repeat(repeatTimes) + row[nameColumn].toString().slice(-1);
      }

      const tmp = [
        row[classColumn],
        row[stdNumberColumn],
        maskedName,
        row[subjectColumn],
        row[sessionColumn],
        row[classroomColumn],
      ];

      modifiedData.push(tmp);
    }
  );

  setRangeValues(bulletinSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  sortBySessionClassroom();
  prettier();
}


function prettier(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bulletinSheet = ss.getSheetByName("公告版補考場次");
  const parametersSheet = ss.getSheetByName("參數區");

  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();

  bulletinSheet.getRange("A1:F1").mergeAcross();
  bulletinSheet.getRange("A1").setValue("高雄高工" + schoolYear + "學年度第" + semester + "學期補考名單");
  bulletinSheet.getRange("A1").setFontSize(20);
  bulletinSheet.getRange(1, 1, bulletinSheet.getMaxRows(), bulletinSheet.getMaxColumns()).setHorizontalAlignment("center");
  bulletinSheet.setFrozenRows(2);
  bulletinSheet.getRange("A2:F").createFilter();
  bulletinSheet.getRange("A2:F").setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
}


function generateRecordSheet() {
  sortBySessionClassroom();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const recordSheet = ss.getSheetByName("試場紀錄表(A表)");

  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const sessionColumn = headers.indexOf("節次");
  const classroomColumn = headers.indexOf("試場");
  const timeColumn = headers.indexOf("時間");
  const classColumn = headers.indexOf("班級");
  const stdNumberColumn = headers.indexOf("學號");
  const nameColumn = headers.indexOf("姓名");
  const subjectColumn = headers.indexOf("科目名稱");
  const classPopulationColumn = headers.indexOf("班級人數");

  // 刪除多餘的欄和列
  recordSheet.clear();
  if (recordSheet.getMaxRows()>5){
    recordSheet.deleteRows(2,recordSheet.getMaxRows() - 5);
  }

  const modifiedData = [
      ["A表："+schoolYear+"學年度第"+semester+"學期補考簽到及違規記錄表    　 　　　                                 監考教師簽名：　　　　　　　　　", "", "", "", "", "", "", "", "", "", "", "",],
      ["節次", "試場",	"時間",	"班級",	"學號",	"姓名",	"科目名稱",	"班級人數",	"考生到考簽名", "違規記錄(打V)", "", "其他違規\n請簡述"],
      ["", "", "", "", "", "", "", "", "", "未帶有照證件", "服儀不整", ""]
    ];

  data.forEach(
    function (row){
      modifiedData.push([
        row[sessionColumn],
        row[classroomColumn],
        row[timeColumn],
        row[classColumn],
        row[stdNumberColumn],
        row[nameColumn],
        row[subjectColumn],
        row[classPopulationColumn],
        "",
        "",
        "",
        ""
      ]);
    }
  );

  setRangeValues(recordSheet.getRange(1, 1, modifiedData.length, modifiedData[0].length), modifiedData);

  // 設定格式美化表格
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

  recordSheet.getRange(2, 1, modifiedData.length + 2, modifiedData[0].length)
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}


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
    smallBagSheet.deleteRows(2,smallBagSheet.getMaxRows()-5);
  }

  const smallBags = [["學年度", "學期", "小袋序號", "節次", "時間", "試場", "班級", "科目名稱", "任課老師", "小袋人數", "電腦", "人工"],];
  const alreadyArranged = [];

  data.forEach(
    function(row){
      if (!alreadyArranged.includes(row[smallBagSerialColumn])){
        const tmp = [
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
    bigBagSheet.deleteRows(2,bigBagSheet.getMaxRows()-5);
  }

  const bigBags = [["學年度", "學期", "大袋序號", "節次", "試場", "補考日期", "時間", "試卷袋序號", "監考教師", "各試場人數"],];
  const alreadyArranged = [];

  const container = {};
  data.forEach(
    function (row){
      if (!Object.keys(container).includes("大袋" + row[bigBagSerialColumn])){
        container["大袋" + row[bigBagSerialColumn]] = [row[smallBagSerialColumn]];
      } else {
        container["大袋" + row[bigBagSerialColumn]].push(row[smallBagSerialColumn]);
      }
    }
  );

  data.forEach(
    function(row){
      if (!alreadyArranged.includes(row[bigBagSerialColumn])){
        const tmp = [
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
