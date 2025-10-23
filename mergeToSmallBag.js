// 參考資料：
// 影片：https://www.youtube.com/watch?v=Bbmrkxk_nbE
// 影片中的程式碼：https://docs.google.com/document/d/1PBMgS3iDPisSOOY0t3h1AZhgvE8M_Uzk2BOz7UywVrE/edit#heading=h.32tuyn8klpk3

function mergeToSmallBag(){
  // 先產生合併列印小袋用的資料
  generateSmallBagData();

  const runtimeCountStart = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const smallBagSheet = ss.getSheetByName("小袋封面套印用資料");
  const [headers,...data] = smallBagSheet.getDataRange().getValues();
  const schoolYear = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const folderId = parametersSheet.getRange("B10").getValue();
  const smallTemplateId = parametersSheet.getRange("B12").getValue();
  const smallTemplate = DriveApp.getFileById(smallTemplateId);
  const makeUpDate = parametersSheet.getRange("B13").getValue();
  const folder = DriveApp.getFolderById(folderId);

  // 學年度	學期	小袋序號	節次	時間	試場	班級	科目名稱	任課老師	小袋人數	電腦	人工
  const schoolYearColumn = headers.indexOf("學年度");
  const semesterColumn = headers.indexOf("學期");
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

  const BATCH_SIZE = 50;  // 每個 PDF 檔的頁數，太大或太小的數字都會減慢整體速度
  const numberOfDigits = data.length.toString().length;

  let mergedFilename = schoolYear + "學年度第" + semester + "學期補考小袋封面";
  let mergedFile = smallTemplate.makeCopy(mergedFilename, folder);
  let mergedDoc = DocumentApp.openById(mergedFile.getId()).setMarginTop(0).setMarginBottom(0);
  let body = mergedDoc.getBody();
  body.removeChild(body.getTables()[0]);

  let tmpDoc = DocumentApp.openById(smallTemplate.makeCopy("暫時的小袋" , folder).getId());
  let tmpBody = tmpDoc.getBody();
  let tmpBodyClone = tmpBody.copy();

  for (let i=0; i < data.length; i++){
    tmpBodyClone.replaceText("«學年度»" ,data[i][schoolYearColumn]);
    tmpBodyClone.replaceText("«學期»" ,data[i][semesterColumn]);
    tmpBodyClone.replaceText("«補考日期»", makeUpDate)
    tmpBodyClone.replaceText("«小袋序號»" ,data[i][smallBagSerialColumn]);
    tmpBodyClone.replaceText("«節次»" ,data[i][sessionColumn]);
    tmpBodyClone.replaceText("«時間»" ,data[i][timeColumn]);
    tmpBodyClone.replaceText("«試場»" ,data[i][classroomColumn]);
    tmpBodyClone.replaceText("«班級»" ,data[i][classColumn]);
    tmpBodyClone.replaceText("«科目名稱»" ,data[i][subjectNameColumn]);
    tmpBodyClone.replaceText("«任課老師»" ,data[i][teacherColumn]);
    tmpBodyClone.replaceText("«班級人數»" ,data[i][smallBagPopulationColumn]);
    tmpBodyClone.replaceText("«電腦»" ,data[i][byComputerColumn]);
    tmpBodyClone.replaceText("«人工»" ,data[i][byHandColumn]);

    const tmpTable = tmpBodyClone.getTables()[0];
    const studentsList = getStudents(data[i][smallBagSerialColumn]);
    const studentsTable = tmpTable.getCell(1,0).setPaddingTop(0).setPaddingBottom(0).appendTable(studentsList);

    const style = {}
    if (studentsList.length > 23){
      style[DocumentApp.Attribute.FONT_SIZE] = 7;
    } else {
      style[DocumentApp.Attribute.FONT_SIZE] = 8;
    }
    studentsTable.setAttributes(style);
    for (let j=0; j < studentsTable.getNumRows() ; j++){
      for (let k=0; k < studentsTable.getRow(j).getNumCells(); k++){
        const cellStyle = {}
        cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
        studentsTable.getRow(j).getCell(k)
          .setPaddingTop(0)
          .setPaddingBottom(0)
          .getChild(0).asParagraph().setAttributes(cellStyle);

        switch (k){
          case 0:  // 編號
            studentsTable.getRow(j).getCell(k).setWidth(20);
            break;
          case 1:  // 班級
            studentsTable.getRow(j).getCell(k).setWidth(60);
            break;
          case 2:  // 學號
            studentsTable.getRow(j).getCell(k).setWidth(60);
            break;
          case 3:  // 姓名
            studentsTable.getRow(j).getCell(k).setWidth(60);
            break;
          case 4:  // 科目名稱
            studentsTable.getRow(j).getCell(k).setWidth(140);
            break;
        }
      }
    }

    body.appendTable(tmpTable.copy());
    if (i % BATCH_SIZE != BATCH_SIZE - 1){
      body.appendPageBreak();
    }
    tmpBodyClone = tmpBody.copy();
    
    if(i % BATCH_SIZE == BATCH_SIZE - 1){
      // 設定字體後存檔
      body.editAsText().setFontFamily("Noto Sans TC");
      mergedDoc.saveAndClose();

      // 轉換成 PDF 檔，且在檔名加上小袋序號的範圍
      const pdfBlob = mergedFile.getAs('application/pdf');
      const pdfName = mergedFilename + "_" + ( 1 + BATCH_SIZE * Math.floor(i / BATCH_SIZE)).toLocaleString('en-US', {minimumIntegerDigits: numberOfDigits, useGrouping:false}) + "-" + (BATCH_SIZE * (1 + Math.floor(i / BATCH_SIZE))).toLocaleString('en-US', {minimumIntegerDigits: numberOfDigits, useGrouping:false})  + ".pdf";
      const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName)
      pdfFile.moveTo(folder);

      // 把舊的 mergedFile 移到垃圾桶後，再重新產生新的 mergedFile
      mergedFile.setTrashed(true);
      mergedFile = smallTemplate.makeCopy(mergedFilename, folder);
      mergedDoc = DocumentApp.openById(mergedFile.getId());
      body = mergedDoc.getBody();
      body.removeChild(body.getTables()[0]);
    }
  }

  // 將暫存檔移到垃圾桶
  tmpDoc.saveAndClose();
  DriveApp.getFileById(tmpDoc.getId()).setTrashed(true);

  if (data.length % BATCH_SIZE != 0){
    // 設定字體為 Noto Sans TC 後存檔
    body.editAsText().setFontFamily("Noto Sans TC");
    mergedDoc.saveAndClose();

    // 把最後一批轉換成 PDF 檔
    const pdfBlob = mergedFile.getAs('application/pdf');
    const pdfName = mergedFilename + "_" + (BATCH_SIZE * Math.floor(data.length / BATCH_SIZE).toLocaleString('en-US', {minimumIntegerDigits: numberOfDigits, useGrouping:false}) + 1) + "-" + data.length + ".pdf";
    const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName)
    pdfFile.moveTo(folder);
  }
  mergedFile.setTrashed(true);

  const newRuntime = runtimeCountStop(runtimeCountStart);
  SpreadsheetApp.getUi().alert("已完成小袋封面合併列印，共" + data.length + "頁，使用" + newRuntime + "秒。");
}


function getStudents(smallBagSerial){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filteredSheet.getDataRange().getValues();
  const smallBagSerialColumn = headers.indexOf("小袋序號");
  const classColumn = headers.indexOf("班級");
  const stdNumberColumn = headers.indexOf("學號");
  const stdNameColumn = headers.indexOf("姓名");
  const subjectNameColumn = headers.indexOf("科目名稱");

  let studentsTable = [["", "班級", "學號", "姓名", "科目", "缺考"]]
  data.forEach(
    function (row){
      if (row[smallBagSerialColumn] == smallBagSerial){
        studentsTable.push(
          [parseInt(studentsTable.length).toString() ,row[classColumn], row[stdNumberColumn], row[stdNameColumn], row[subjectNameColumn], ""]
        );
      }
    }
  );

  return studentsTable;
}
