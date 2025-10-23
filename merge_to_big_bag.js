function mergeToBigBag(){
  // 先產生合併大袋用的資料
  // generateBigBagData();

  const runtimeCountStart = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const bigBagSheet = ss.getSheetByName("大袋封面套印用資料");
  const [headers,...data] = bigBagSheet.getDataRange().getValues();
  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const folderId = parametersSheet.getRange("B10").getValue();
  const bigTemplateId = parametersSheet.getRange("B11").getValue();
  const bigTemplate = DriveApp.getFileById(bigTemplateId);
  const makeUpDate = parametersSheet.getRange("B13").getValue();
  const folder = DriveApp.getFolderById(folderId);

  // 學年度	學期	小袋序號	節次	時間	試場	班級	科目名稱	任課老師	小袋人數	電腦	人工
  const schoolYearColumn = headers.indexOf("學年度");
  const semesterColumn = headers.indexOf("學期");
  const bigBagSerialColumn = headers.indexOf("大袋序號");
  const sessionColumn = headers.indexOf("節次");
  const classroomColumn = headers.indexOf("試場");
  const timeColumn = headers.indexOf("時間");
  const smallBagColumn = headers.indexOf("試卷袋序號");
  const teacherColumn = headers.indexOf("監考教師");
  const bigBagPopulationColumn = headers.indexOf("各試場人數");

  const BATCH_SIZE = 50;  // 每個 PDF 檔的頁數，太大或太小的數字都會減慢整體速度
  const numberOfDigits = data.length.toString().length;

  let mergedFilename = schoolYear + "學年度第" + semester + "學期補考大袋封面";
  let mergedFile = bigTemplate.makeCopy(mergedFilename, folder);
  let mergedDoc = DocumentApp.openById(mergedFile.getId());
  let mergedBody = mergedDoc.getBody().clear();

  let tmpFile = bigTemplate.makeCopy("暫時的大袋" , folder);
  let tmpDoc = DocumentApp.openById(tmpFile.getId());
  let tmpBody = tmpDoc.getBody();
  let tmpBodyClone = tmpBody.copy();
  let numberOfListItems = tmpBody.getListItems().length;

  for (let i=0; i < data.length; i++){
    tmpBodyClone.replaceText("«學年度»" ,data[i][schoolYearColumn]);
    tmpBodyClone.replaceText("«學期»" ,data[i][semesterColumn]);
    tmpBodyClone.replaceText("«大袋序號»" ,data[i][bigBagSerialColumn]);
    tmpBodyClone.replaceText("«節次»" ,data[i][sessionColumn]);
    tmpBodyClone.replaceText("«試場»" ,data[i][classroomColumn]);
    tmpBodyClone.replaceText("«補考日期»", makeUpDate)
    tmpBodyClone.replaceText("«時間»" ,data[i][timeColumn]);
    tmpBodyClone.replaceText("«試卷袋序號»" ,data[i][smallBagColumn]);
    tmpBodyClone.replaceText("«監考教師»" ,data[i][teacherColumn]);
    tmpBodyClone.replaceText("«各試場人數»" ,data[i][bigBagPopulationColumn]);

    mergedBody.appendParagraph(tmpBodyClone.getParagraphs()[0].copy());
    mergedBody.appendParagraph(tmpBodyClone.getParagraphs()[1].copy());
    mergedBody.appendTable(tmpBodyClone.getTables()[0].copy());

    // 將最後幾個 ListItem 重新編碼
    let newLists = mergedBody.getListItems();
    let tmpList = mergedBody.appendListItem("暫");
    for (let j=newLists.length - 1; j > newLists.length - numberOfListItems - 1; j--){
      newLists[j].setListId(tmpList);
    }
    mergedBody.appendParagraph("");
    tmpList.removeFromParent();

    if (i % BATCH_SIZE != BATCH_SIZE - 1){
      mergedBody.appendPageBreak();
    }

    // 建立新的 tmp file
    tmpBodyClone = tmpBody.copy();
    
    if(i % BATCH_SIZE == BATCH_SIZE - 1){
      // 設定字體後存檔
      mergedBody.editAsText().setFontFamily("Noto Sans TC");
      mergedDoc.saveAndClose();

      // 轉換成 PDF 檔，且在檔名加上小袋序號的範圍
      const pdfBlob = mergedFile.getAs('application/pdf');
      const pdfName = mergedFilename + "_" + ( 1 + BATCH_SIZE * Math.floor(i / BATCH_SIZE)).toLocaleString('en-US', {minimumIntegerDigits: numberOfDigits, useGrouping:false}) + "-" + (BATCH_SIZE * (1 + Math.floor(i / BATCH_SIZE))).toLocaleString('en-US', {minimumIntegerDigits: numberOfDigits, useGrouping:false})  + ".pdf";
      const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName)
      pdfFile.moveTo(folder);

      // 把舊的 mergedFile 移到垃圾桶後，再重新產生新的 mergedFile
      mergedFile.setTrashed(true);

      // 不是最後一筆，後面還有其他筆資料時，才產生新的檔案
      if (i != data.length - 1){
        mergedFile = bigTemplate.makeCopy(mergedFilename, folder);
        mergedDoc = DocumentApp.openById(mergedFile.getId());
        mergedBody = mergedDoc.getBody().clear();
      }
    }
  }

  // 將暫存檔移到垃圾桶
  tmpDoc.saveAndClose();
  DriveApp.getFileById(tmpDoc.getId()).setTrashed(true);

  // 真的有畸零筆數所成的檔案才存檔
  if (data.length % BATCH_SIZE != 0){
    // 設定字體為 Noto Sans TC 後存檔
    mergedBody.editAsText().setFontFamily("Noto Sans TC");
    mergedDoc.saveAndClose();

    // 把最後一批轉換成 PDF 檔
    const pdfBlob = mergedFile.getAs('application/pdf');
    const pdfName = mergedFilename + "_" + (BATCH_SIZE * Math.floor(data.length / BATCH_SIZE).toLocaleString('en-US', {minimumIntegerDigits: numberOfDigits, useGrouping:false}) + 1) + "-" + data.length + ".pdf";
    const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName)
    pdfFile.moveTo(folder);
  }
  mergedFile.setTrashed(true);

  const newRuntime = runtimeCountStop(runtimeCountStart);
  SpreadsheetApp.getUi().alert("已完成大袋封面合併列印，共" + data.length + "頁，使用" + newRuntime + "秒。");
}
