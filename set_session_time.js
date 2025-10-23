function setSessionTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const sessionColumn = headers.indexOf("節次");
  const timeColumn = headers.indexOf("時間");

  const timeTableSheet = ss.getSheetByName("節次時間表");
  const [timeHeaders, ...timeData] = timeTableSheet.getDataRange().getValues();

  const timeTable = {}
  timeData.forEach(
    function(timeRow){
      timeTable[timeRow[0]] = timeRow[1];
    }
  );

  let modifiedData = data.map(
    function(row){
      row[timeColumn] = timeTable[row[sessionColumn]];
      return row;
    }
  );

  if(modifiedData.length == data.length){
    setRangeValues(filteredSheet.getRange(2, 1, modifiedData.length, modifiedData[0].length), modifiedData);
  } else {
    Logger.log("寫入節次時間失敗！");
    SpreadsheetApp.getUi().alert("寫入節次時間失敗！");
  }
}
