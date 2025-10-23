function merge_to_big_bag(){
  // 先產生合併大袋用的資料
  // generate_big_bag_data();

  const runtime_count_start = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const big_bag_sheet = ss.getSheetByName("大袋封面套印用資料");
  const [headers,...data] = big_bag_sheet.getDataRange().getValues();
  const school_year = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const folder_id = parametersSheet.getRange("B10").getValue();
  const big_template_id = parametersSheet.getRange("B11").getValue();
  const big_template = DriveApp.getFileById(big_template_id);
  const make_up_date = parametersSheet.getRange("B13").getValue();
  const folder = DriveApp.getFolderById(folder_id);

  // 學年度	學期	小袋序號	節次	時間	試場	班級	科目名稱	任課老師	小袋人數	電腦	人工
  const school_year_column = headers.indexOf("學年度");
  const semester_column = headers.indexOf("學期");
  const big_bag_serial_column = headers.indexOf("大袋序號");
  const session_column = headers.indexOf("節次");
  const classroom_column = headers.indexOf("試場");
  const time_column = headers.indexOf("時間");
  const small_bag_column = headers.indexOf("試卷袋序號");
  const teacher_column = headers.indexOf("監考教師");
  const big_bag_population_column = headers.indexOf("各試場人數");

  const BATCH_SIZE = 50;  // 每個 PDF 檔的頁數，太大或太小的數字都會減慢整體速度
  const number_of_digits = data.length.toString().length;

  let merged_filename = school_year + "學年度第" + semester + "學期補考大袋封面";
  let merged_file = big_template.makeCopy(merged_filename, folder);
  let merged_doc = DocumentApp.openById(merged_file.getId());
  let merged_body = merged_doc.getBody().clear();

  let tmp_file = big_template.makeCopy("暫時的大袋" , folder);
  let tmp_doc = DocumentApp.openById(tmp_file.getId());
  let tmp_body = tmp_doc.getBody();
  let tmp_body_1 = tmp_body.copy();
  let number_of_listitems = tmp_body.getListItems().length;

  for (i=0; i < data.length; i++){
    tmp_body_1.replaceText("«學年度»" ,data[i][school_year_column]);
    tmp_body_1.replaceText("«學期»" ,data[i][semester_column]);
    tmp_body_1.replaceText("«大袋序號»" ,data[i][big_bag_serial_column]);
    tmp_body_1.replaceText("«節次»" ,data[i][session_column]);
    tmp_body_1.replaceText("«試場»" ,data[i][classroom_column]);
    tmp_body_1.replaceText("«補考日期»", make_up_date)
    tmp_body_1.replaceText("«時間»" ,data[i][time_column]);
    tmp_body_1.replaceText("«試卷袋序號»" ,data[i][small_bag_column]);
    tmp_body_1.replaceText("«監考教師»" ,data[i][teacher_column]);
    tmp_body_1.replaceText("«各試場人數»" ,data[i][big_bag_population_column]);

    merged_body.appendParagraph(tmp_body_1.getParagraphs()[0].copy());
    merged_body.appendParagraph(tmp_body_1.getParagraphs()[1].copy());
    merged_body.appendTable(tmp_body_1.getTables()[0].copy());

    // 將最後幾個 ListItem 重新編碼
    let new_lists = merged_body.getListItems();
    let tmp_list = merged_body.appendListItem("暫");
    for (let j=new_lists.length - 1; j > new_lists.length - number_of_listitems - 1; j--){
      new_lists[j].setListId(tmp_list);
    }
    merged_body.appendParagraph("");
    tmp_list.removeFromParent();

    if (i % BATCH_SIZE != BATCH_SIZE - 1){
      merged_body.appendPageBreak();
    }

    // 建立新的 tmp file
    tmp_body_1 = tmp_body.copy();
    
    if(i % BATCH_SIZE == BATCH_SIZE - 1){
      // 設定字體後存檔
      merged_body.editAsText().setFontFamily("Noto Sans TC");
      merged_doc.saveAndClose();

      // 轉換成 PDF 檔，且在檔名加上小袋序號的範圍
      let pdf_blob = merged_file.getAs('application/pdf');
      let pdf_name = merged_filename + "_" + ( 1 + BATCH_SIZE * Math.floor(i / BATCH_SIZE)).toLocaleString('en-US', {minimumIntegerDigits: number_of_digits, useGrouping:false}) + "-" + (BATCH_SIZE * (1 + Math.floor(i / BATCH_SIZE))).toLocaleString('en-US', {minimumIntegerDigits: number_of_digits, useGrouping:false})  + ".pdf";
      let pdf_file = DriveApp.createFile(pdf_blob).setName(pdf_name)
      pdf_file.moveTo(folder);

      // 把舊的 merged_file 移到垃圾桶後，再重新產生新的 merged_file
      merged_file.setTrashed(true);

      // 不是最後一筆，後面還有其他筆資料時，才產生新的檔案
      if (i != data.length - 1){
        merged_file = big_template.makeCopy(merged_filename, folder);
        merged_doc = DocumentApp.openById(merged_file.getId());
        merged_body = merged_doc.getBody().clear();
      }
    }
  }

  // 將暫存檔移到垃圾桶
  tmp_doc.saveAndClose();
  DriveApp.getFileById(tmp_doc.getId()).setTrashed(true);

  // 真的有畸零筆數所成的檔案才存檔
  if (data.length % BATCH_SIZE != 0){
    // 設定字體為 Noto Sans TC 後存檔
    merged_body.editAsText().setFontFamily("Noto Sans TC");
    merged_doc.saveAndClose();

    // 把最後一批轉換成 PDF 檔
    pdf_blob = merged_file.getAs('application/pdf');
    pdf_name = merged_filename + "_" + (BATCH_SIZE * Math.floor(data.length / BATCH_SIZE).toLocaleString('en-US', {minimumIntegerDigits: number_of_digits, useGrouping:false}) + 1) + "-" + data.length + ".pdf";
    pdf_file = DriveApp.createFile(pdf_blob).setName(pdf_name)
    pdf_file.moveTo(folder);
  }
  merged_file.setTrashed(true);

  newRuntime = runtime_count_stop(runtime_count_start);
  SpreadsheetApp.getUi().alert("已完成大袋封面合併列印，共" + data.length + "頁，使用" + newRuntime + "秒。");
}
