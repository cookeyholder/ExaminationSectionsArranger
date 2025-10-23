function unfilted_code_complete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const unfiltered_sheet = ss.getSheetByName("註冊組補考名單");
  const [unfiltered_sheetHeaders, ...unfiltered_data] = unfiltered_sheet.getDataRange().getValues();

  const class_column = unfiltered_sheetHeaders.indexOf("班級");
  const subject_column = unfiltered_sheetHeaders.indexOf("科目");
  const code_column = unfiltered_sheetHeaders.indexOf("科目代碼補完");
  const subject_name_column = unfiltered_sheetHeaders.indexOf("科目名稱");

  const department_to_group = {
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

  const grade_to_year = {
    "一": parseInt(parametersSheet.getRange("B2").getValue()),
    "二": parseInt(parametersSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parametersSheet.getRange("B2").getValue()) - 2,
  };

  let modified_data = [];
  unfiltered_data.forEach(
    function(row){
      let tmp = row[subject_column].toString().split(".")[0];
      if(tmp.length == 16){
        row[code_column] = tmp.slice(0, 3) + "553401" + tmp.slice(3, 9) + "0" + tmp.slice(9);
      } else {
        row[code_column] = grade_to_year[row[class_column].toString().slice(2, 3)] + "553401V" + department_to_group[tmp.slice(0,3)] + tmp.slice(0,3) + "0" + tmp.slice(3);
      }

      row[subject_name_column] = row[subject_column].toString().split(".")[1]
      modified_data.push(row);
    }
  );

  if(modified_data.length == unfiltered_data.length){
    set_range_values(unfiltered_sheet.getRange(2, 1, modified_data.length, modified_data[0].length), modified_data);
  } else {
    Logger.log("課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("課程代碼補完失敗！");
  }
}



function open_code_complete() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const open_sheet = ss.getSheetByName("開課資料(查詢任課教師用)");
  const [open_sheetHeaders, ...open_data] = open_sheet.getDataRange().getValues();

  const class_column = open_sheetHeaders.indexOf("班級名稱");
  const code_column = open_sheetHeaders.indexOf("科目代碼");
  const complete_column = open_sheetHeaders.indexOf("科目代碼補完");
  const subject_name_column = open_sheetHeaders.indexOf("科目名稱");

  const department_to_group = {
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

  const grade_to_year = {
    "一": parseInt(parametersSheet.getRange("B2").getValue()),
    "二": parseInt(parametersSheet.getRange("B2").getValue()) - 1,
    "三": parseInt(parametersSheet.getRange("B2").getValue()) - 2,
  };

  let modified_data = [];
  open_data.forEach(
    function(row){
      let tmp = row[code_column];
      if(row[code_column].length == 16){
        row[complete_column] = tmp.slice(0, 3) + "553401" + tmp.slice(3, 9) + "0" + tmp.slice(9);
      } else {
        row[complete_column] = grade_to_year[row[class_column].toString().slice(2, 3)] + "553401V" + department_to_group[tmp.slice(0,3)] + tmp.slice(0,3) + "0" + tmp.slice(3);
      }


      modified_data.push(row);
    }
  );

  if(modified_data.length == open_data.length){
    set_range_values(open_sheet.getRange(2, 1, modified_data.length, modified_data[0].length), modified_data);
  } else {
    Logger.log("開課資料課程代碼補完失敗！");
    SpreadsheetApp.getUi().alert("開課資料課程代碼補完失敗！");
  }
}
