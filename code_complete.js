function unfilteredCodeComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const unfilteredSheet = ss.getSheetByName("註冊組補考名單");
  const [unfilteredSheetHeaders, ...unfilteredData] = unfilteredSheet.getDataRange().getValues();

  const classColumn = unfilteredSheetHeaders.indexOf("班級");
  const subjectColumn = unfilteredSheetHeaders.indexOf("科目");
  const codeColumn = unfilteredSheetHeaders.indexOf("科目代碼補完");
  const subjectNameColumn = unfilteredSheetHeaders.indexOf("科目名稱");

  const departmentToGroup = {
    "301": "21",
    "303": "22",
    "305": "23",
    "306": "23",
    "308": "23",
    "309": "23",
    "311": "25",
    "315": "24",
    "373": "28",
    "374": "21",
  };

  const gradeToYear = {
    "一": parseInt(parametersSheet.getRange("B2").getValue()),
    "二": parseInt(parametersSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parametersSheet.getRange("B2").getValue()) - 2,
  };

  let modifiedData = [];
  unfilteredData.forEach(
    function(row){
      let tmp = row[subjectColumn].toString().split(".")[0];
      if(tmp.length == 16){
        row[codeColumn] = tmp.slice(0, 3) + "553401" + tmp.slice(3, 9) + "0" + tmp.slice(9);
      } else {
        row[codeColumn] = gradeToYear[row[classColumn].toString().slice(2, 3)] + "553401V" + departmentToGroup[tmp.slice(0,3)] + tmp.slice(0,3) + "0" + tmp.slice(3);
      }

      row[subjectNameColumn] = row[subjectColumn].toString().split(".")[1]
      modifiedData.push(row);
    }
  );

  if(modifiedData.length == unfilteredData.length){
    setRangeValues(unfilteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("課程代碼補完失敗！");
  }
}



function openCodeComplete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const openSheet = ss.getSheetByName("開課資料(查詢任課教師用)");
  const [openSheetHeaders, ...openData] = openSheet.getDataRange().getValues();

  const classColumn = openSheetHeaders.indexOf("班級名稱");
  const codeColumn = openSheetHeaders.indexOf("科目代碼");
  const completeColumn = openSheetHeaders.indexOf("科目代碼補完");
  const subjectNameColumn = openSheetHeaders.indexOf("科目名稱");

  const departmentToGroup = {
    "301": "21",
    "303": "22",
    "305": "23",
    "306": "23",
    "308": "23",
    "309": "23",
    "311": "25",
    "315": "24",
    "373": "28",
    "374": "21",
  };

  const gradeToYear = {
    "一": parseInt(parametersSheet.getRange("B2").getValue()),
    "二": parseInt(parametersSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parametersSheet.getRange("B2").getValue()) - 2,
  };

  let modifiedData = [];
  openData.forEach(
    function(row){
      let tmp = row[codeColumn];
      if(row[codeColumn].length == 16){
        row[completeColumn] = tmp.slice(0, 3) + "553401" + tmp.slice(3, 9) + "0" + tmp.slice(9);
      } else {
        row[completeColumn] = gradeToYear[row[classColumn].toString().slice(2, 3)] + "553401V" + departmentToGroup[tmp.slice(0,3)] + tmp.slice(0,3) + "0" + tmp.slice(3);
      }


      modifiedData.push(row);
    }
  );

  if(modifiedData.length == openData.length){
    setRangeValues(openSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("開課資料課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("開課資料課程代碼補完失敗！");
  }
}
