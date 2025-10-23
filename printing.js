function mergeSmallBagPdfFiles(){
  composeSmallBagDataset();

  const runtimeStart = new Date();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const smallBagSheet = spreadsheet.getSheetByName("小袋封面套印用資料");
  const [headerRow, ...smallBagRows] = smallBagSheet.getDataRange().getValues();
  const schoolYearValue = parameterSheet.getRange("B2").getValue();
  const semesterValue = parameterSheet.getRange("B3").getValue();
  const destinationFolderId = parameterSheet.getRange("B10").getValue();
  const templateId = parameterSheet.getRange("B12").getValue();
  const templateFile = DriveApp.getFileById(templateId);
  const examDateValue = parameterSheet.getRange("B13").getValue();
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);

  const schoolYearIndex = headerRow.indexOf("學年度");
  const semesterIndex = headerRow.indexOf("學期");
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

  const batchSize = 50;
  const digitCount = smallBagRows.length.toString().length;

  let mergedFileName = schoolYearValue + "學年度第" + semesterValue + "學期補考小袋封面";
  let mergedDocFile = templateFile.makeCopy(mergedFileName, destinationFolder);
  let mergedDocument = DocumentApp.openById(mergedDocFile.getId()).setMarginTop(0).setMarginBottom(0);
  let mergedBody = mergedDocument.getBody();
  mergedBody.removeChild(mergedBody.getTables()[0]);

  let temporaryDocument = DocumentApp.openById(templateFile.makeCopy("暫時的小袋" , destinationFolder).getId());
  let temporaryBody = temporaryDocument.getBody();
  let temporaryBodyClone = temporaryBody.copy();

  for (let rowIndex = 0; rowIndex < smallBagRows.length; rowIndex++){
    temporaryBodyClone.replaceText("«學年度»" ,smallBagRows[rowIndex][schoolYearIndex]);
    temporaryBodyClone.replaceText("«學期»" ,smallBagRows[rowIndex][semesterIndex]);
    temporaryBodyClone.replaceText("«補考日期»", examDateValue);
    temporaryBodyClone.replaceText("«小袋序號»" ,smallBagRows[rowIndex][smallBagIndex]);
    temporaryBodyClone.replaceText("«節次»" ,smallBagRows[rowIndex][sessionIndex]);
    temporaryBodyClone.replaceText("«時間»" ,smallBagRows[rowIndex][timeIndex]);
    temporaryBodyClone.replaceText("«試場»" ,smallBagRows[rowIndex][roomIndex]);
    temporaryBodyClone.replaceText("«班級»" ,smallBagRows[rowIndex][classIndex]);
    temporaryBodyClone.replaceText("«科目名稱»" ,smallBagRows[rowIndex][subjectIndex]);
    temporaryBodyClone.replaceText("«任課老師»" ,smallBagRows[rowIndex][teacherIndex]);
    temporaryBodyClone.replaceText("«班級人數»" ,smallBagRows[rowIndex][smallBagPopulationIndex]);
    temporaryBodyClone.replaceText("«電腦»" ,smallBagRows[rowIndex][computerIndex]);
    temporaryBodyClone.replaceText("«人工»" ,smallBagRows[rowIndex][manualIndex]);

    const templateTable = temporaryBodyClone.getTables()[0];
    const studentListing = buildSmallBagStudentTable(smallBagRows[rowIndex][smallBagIndex]);
    const studentTable = templateTable.getCell(1,0).setPaddingTop(0).setPaddingBottom(0).appendTable(studentListing);

    const tableStyle = {};
    tableStyle[DocumentApp.Attribute.FONT_SIZE] = studentListing.length > 23 ? 7 : 8;
    studentTable.setAttributes(tableStyle);
    for (let studentRowIndex = 0; studentRowIndex < studentTable.getNumRows() ; studentRowIndex++){
      for (let studentCellIndex = 0; studentCellIndex < studentTable.getRow(studentRowIndex).getNumCells(); studentCellIndex++){
        const cellStyle = {};
        cellStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
        studentTable.getRow(studentRowIndex).getCell(studentCellIndex)
          .setPaddingTop(0)
          .setPaddingBottom(0)
          .getChild(0).asParagraph().setAttributes(cellStyle);

        switch (studentCellIndex){
          case 0:
            studentTable.getRow(studentRowIndex).getCell(studentCellIndex).setWidth(20);
            break;
          case 1:
          case 2:
          case 3:
            studentTable.getRow(studentRowIndex).getCell(studentCellIndex).setWidth(60);
            break;
          case 4:
            studentTable.getRow(studentRowIndex).getCell(studentCellIndex).setWidth(140);
            break;
        }
      }
    }

    mergedBody.appendTable(templateTable.copy());
    if (rowIndex % batchSize !== batchSize - 1){
      mergedBody.appendPageBreak();
    }
    temporaryBodyClone = temporaryBody.copy();
    
    if(rowIndex % batchSize === batchSize - 1){
      mergedBody.editAsText().setFontFamily("Noto Sans TC");
      mergedDocument.saveAndClose();

      const pdfBlob = mergedDocFile.getAs('application/pdf');
      const pdfName =
        mergedFileName + "_" +
        ( 1 + batchSize * Math.floor(rowIndex / batchSize)).toLocaleString('en-US', {minimumIntegerDigits: digitCount, useGrouping:false}) +
        "-" +
        (batchSize * (1 + Math.floor(rowIndex / batchSize))).toLocaleString('en-US', {minimumIntegerDigits: digitCount, useGrouping:false}) +
        ".pdf";
      const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName);
      pdfFile.moveTo(destinationFolder);

      mergedDocFile.setTrashed(true);
      mergedDocFile = templateFile.makeCopy(mergedFileName, destinationFolder);
      mergedDocument = DocumentApp.openById(mergedDocFile.getId()).setMarginTop(0).setMarginBottom(0);
      mergedBody = mergedDocument.getBody();
      mergedBody.removeChild(mergedBody.getTables()[0]);
    }
  }

  temporaryDocument.saveAndClose();
  DriveApp.getFileById(temporaryDocument.getId()).setTrashed(true);

  if (smallBagRows.length % batchSize !== 0){
    mergedBody.editAsText().setFontFamily("Noto Sans TC");
    mergedDocument.saveAndClose();

    const pdfBlob = mergedDocFile.getAs('application/pdf');
    const pdfName =
      mergedFileName + "_" +
      (batchSize * Math.floor(smallBagRows.length / batchSize).toLocaleString('en-US', {minimumIntegerDigits: digitCount, useGrouping:false}) + 1) +
      "-" +
      smallBagRows.length +
      ".pdf";
    const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName);
    pdfFile.moveTo(destinationFolder);
  }
  mergedDocFile.setTrashed(true);

  const elapsedSeconds = calculateElapsedSeconds(runtimeStart);
  SpreadsheetApp.getUi().alert("已完成小袋封面合併列印，共" + smallBagRows.length + "頁，使用" + elapsedSeconds + "秒。");
}


function buildSmallBagStudentTable(smallBagNumber){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = spreadsheet.getSheetByName("排入考程的補考名單");
  const [headerRow, ...candidateRows] = filteredSheet.getDataRange().getValues();
  const smallBagIndex = headerRow.indexOf("小袋序號");
  const classIndex = headerRow.indexOf("班級");
  const studentNumberIndex = headerRow.indexOf("學號");
  const studentNameIndex = headerRow.indexOf("姓名");
  const subjectIndex = headerRow.indexOf("科目名稱");

  const studentTable = [["", "班級", "學號", "姓名", "科目", "缺考"]];
  candidateRows.forEach(
    function(examineeRow){
      if (examineeRow[smallBagIndex] === smallBagNumber){
        studentTable.push(
          [parseInt(studentTable.length).toString(), examineeRow[classIndex], examineeRow[studentNumberIndex], examineeRow[studentNameIndex], examineeRow[subjectIndex], ""]
        );
      }
    }
  );

  return studentTable;
}


function mergeBigBagPdfFiles(){
  // composeBigBagDataset(); // 保留人工設定監考教師的彈性，必要時再手動呼叫。

  const runtimeStart = new Date();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const parameterSheet = spreadsheet.getSheetByName("參數區");
  const bigBagSheet = spreadsheet.getSheetByName("大袋封面套印用資料");
  const [headerRow, ...bigBagRows] = bigBagSheet.getDataRange().getValues();
  const schoolYearValue = parameterSheet.getRange("B2").getValue();
  const semesterValue = parameterSheet.getRange("B3").getValue();
  const destinationFolderId = parameterSheet.getRange("B10").getValue();
  const templateId = parameterSheet.getRange("B11").getValue();
  const templateFile = DriveApp.getFileById(templateId);
  const examDateValue = parameterSheet.getRange("B13").getValue();
  const destinationFolder = DriveApp.getFolderById(destinationFolderId);

  const schoolYearIndex = headerRow.indexOf("學年度");
  const semesterIndex = headerRow.indexOf("學期");
  const bigBagIndex = headerRow.indexOf("大袋序號");
  const sessionIndex = headerRow.indexOf("節次");
  const roomIndex = headerRow.indexOf("試場");
  const timeIndex = headerRow.indexOf("時間");
  const paperBagRangeIndex = headerRow.indexOf("試卷袋序號");
  const invigilatorIndex = headerRow.indexOf("監考教師");
  const populationIndex = headerRow.indexOf("各試場人數");

  const batchSize = 50;
  const digitCount = bigBagRows.length.toString().length;

  let mergedFileName = schoolYearValue + "學年度第" + semesterValue + "學期補考大袋封面";
  let mergedDocFile = templateFile.makeCopy(mergedFileName, destinationFolder);
  let mergedDocument = DocumentApp.openById(mergedDocFile.getId());
  let mergedBody = mergedDocument.getBody().clear();

  let temporaryFile = templateFile.makeCopy("暫時的大袋" , destinationFolder);
  let temporaryDocument = DocumentApp.openById(temporaryFile.getId());
  let temporaryBody = temporaryDocument.getBody();
  let temporaryBodyClone = temporaryBody.copy();
  const listItemCount = temporaryBody.getListItems().length;

  for (let rowIndex = 0; rowIndex < bigBagRows.length; rowIndex++){
    temporaryBodyClone.replaceText("«學年度»" ,bigBagRows[rowIndex][schoolYearIndex]);
    temporaryBodyClone.replaceText("«學期»" ,bigBagRows[rowIndex][semesterIndex]);
    temporaryBodyClone.replaceText("«大袋序號»" ,bigBagRows[rowIndex][bigBagIndex]);
    temporaryBodyClone.replaceText("«節次»" ,bigBagRows[rowIndex][sessionIndex]);
    temporaryBodyClone.replaceText("«試場»" ,bigBagRows[rowIndex][roomIndex]);
    temporaryBodyClone.replaceText("«補考日期»", examDateValue);
    temporaryBodyClone.replaceText("«時間»" ,bigBagRows[rowIndex][timeIndex]);
    temporaryBodyClone.replaceText("«試卷袋序號»" ,bigBagRows[rowIndex][paperBagRangeIndex]);
    temporaryBodyClone.replaceText("«監考教師»" ,bigBagRows[rowIndex][invigilatorIndex]);
    temporaryBodyClone.replaceText("«各試場人數»" ,bigBagRows[rowIndex][populationIndex]);

    mergedBody.appendParagraph(temporaryBodyClone.getParagraphs()[0].copy());
    mergedBody.appendParagraph(temporaryBodyClone.getParagraphs()[1].copy());
    mergedBody.appendTable(temporaryBodyClone.getTables()[0].copy());

    const listItems = mergedBody.getListItems();
    const placeholderList = mergedBody.appendListItem("暫");
    for (let listIndex = listItems.length - 1; listIndex > listItems.length - listItemCount - 1; listIndex--){
      listItems[listIndex].setListId(placeholderList);
    }
    mergedBody.appendParagraph("");
    placeholderList.removeFromParent();

    if (rowIndex % batchSize !== batchSize - 1){
      mergedBody.appendPageBreak();
    }

    temporaryBodyClone = temporaryBody.copy();
    
    if(rowIndex % batchSize === batchSize - 1){
      mergedBody.editAsText().setFontFamily("Noto Sans TC");
      mergedDocument.saveAndClose();

      const pdfBlob = mergedDocFile.getAs('application/pdf');
      const pdfName =
        mergedFileName + "_" +
        ( 1 + batchSize * Math.floor(rowIndex / batchSize)).toLocaleString('en-US', {minimumIntegerDigits: digitCount, useGrouping:false}) +
        "-" +
        (batchSize * (1 + Math.floor(rowIndex / batchSize))).toLocaleString('en-US', {minimumIntegerDigits: digitCount, useGrouping:false}) +
        ".pdf";
      const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName);
      pdfFile.moveTo(destinationFolder);

      mergedDocFile.setTrashed(true);

      if (rowIndex !== bigBagRows.length - 1){
        mergedDocFile = templateFile.makeCopy(mergedFileName, destinationFolder);
        mergedDocument = DocumentApp.openById(mergedDocFile.getId());
        mergedBody = mergedDocument.getBody().clear();
      }
    }
  }

  temporaryDocument.saveAndClose();
  DriveApp.getFileById(temporaryDocument.getId()).setTrashed(true);

  if (bigBagRows.length % batchSize !== 0){
    mergedBody.editAsText().setFontFamily("Noto Sans TC");
    mergedDocument.saveAndClose();

    const pdfBlob = mergedDocFile.getAs('application/pdf');
    const pdfName =
      mergedFileName + "_" +
      (batchSize * Math.floor(bigBagRows.length / batchSize).toLocaleString('en-US', {minimumIntegerDigits: digitCount, useGrouping:false}) + 1) +
      "-" +
      bigBagRows.length +
      ".pdf";
    const pdfFile = DriveApp.createFile(pdfBlob).setName(pdfName);
    pdfFile.moveTo(destinationFolder);
  }
  mergedDocFile.setTrashed(true);

  const elapsedSeconds = calculateElapsedSeconds(runtimeStart);
  SpreadsheetApp.getUi().alert("已完成大袋封面合併列印，共" + bigBagRows.length + "頁，使用" + elapsedSeconds + "秒。");
}
