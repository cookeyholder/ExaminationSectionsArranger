// 參考資料：
// 影片：https://www.youtube.com/watch?v=Bbmrkxk_nbE
// 影片中的程式碼：https://docs.google.com/document/d/1PBMgS3iDPisSOOY0t3h1AZhgvE8M_Uzk2BOz7UywVrE/edit#heading=h.32tuyn8klpk3


function merge_to_small_bag(){
  // 先產生合併列印小袋用的資料
  generate_small_bag_data();

  const runtime_count_start = new Date();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const parametersSheet = ss.getSheetByName("參數區");
  const small_bag_sheet = ss.getSheetByName("小袋封面套印用資料");
  const [headers,...data] = small_bag_sheet.getDataRange().getValues();
  const school_year = parametersSheet.getRange("B2").getValue();
  const semester = parametersSheet.getRange("B3").getValue();
  const folder_id = parametersSheet.getRange("B10").getValue();
  const small_template_id = parametersSheet.getRange("B12").getValue();
  const small_template = DriveApp.getFileById(small_template_id);
  const make_up_date = parametersSheet.getRange("B13").getValue();
  const folder = DriveApp.getFolderById(folder_id);

  // 學年度	學期	小袋序號	節次	時間	試場	班級	科目名稱	任課老師	小袋人數	電腦	人工
  const school_year_column = headers.indexOf("學年度");
  const semester_column = headers.indexOf("學期");
  const small_bag_serial_column = headers.indexOf("小袋序號");
  const session_column = headers.indexOf("節次");
  const time_column = headers.indexOf("時間");
  const classroom_column = headers.indexOf("試場");
  const class_column = headers.indexOf("班級");
  const subject_name_column = headers.indexOf("科目名稱");
  const teacher_column = headers.indexOf("任課老師");
  const small_bag_population_column = headers.indexOf("小袋人數");
  const bycomputer_column = headers.indexOf("電腦");
  const byhand_column = headers.indexOf("人工");

  const BATCH_SIZE = 50;  // 每個 PDF 檔的頁數，太大或太小的數字都會減慢整體速度
  const number_of_digits = data.length.toString().length;

  let merged_filename = school_year + "學年度第" + semester + "學期補考小袋封面";
  let merged_file = small_template.makeCopy(merged_filename, folder);
  let merged_doc = DocumentApp.openById(merged_file.getId()).setMarginTop(0).setMarginBottom(0);
  let body = merged_doc.getBody();
  body.removeChild(body.getTables()[0]);

  let tmp_doc = DocumentApp.openById(small_template.makeCopy("暫時的小袋" , folder).getId());
  let tmp_body = tmp_doc.getBody();
  let tmp_body_1 = tmp_body.copy();

  for (i=0; i < data.length; i++){
    tmp_body_1.replaceText("«學年度»" ,data[i][school_year_column]);
    tmp_body_1.replaceText("«學期»" ,data[i][semester_column]);
    tmp_body_1.replaceText("«補考日期»", make_up_date)
    tmp_body_1.replaceText("«小袋序號»" ,data[i][small_bag_serial_column]);
    tmp_body_1.replaceText("«節次»" ,data[i][session_column]);
    tmp_body_1.replaceText("«時間»" ,data[i][time_column]);
    tmp_body_1.replaceText("«試場»" ,data[i][classroom_column]);
    tmp_body_1.replaceText("«班級»" ,data[i][class_column]);
    tmp_body_1.replaceText("«科目名稱»" ,data[i][subject_name_column]);
    tmp_body_1.replaceText("«任課老師»" ,data[i][teacher_column]);
    tmp_body_1.replaceText("«班級人數»" ,data[i][small_bag_population_column]);
    tmp_body_1.replaceText("«電腦»" ,data[i][bycomputer_column]);
    tmp_body_1.replaceText("«人工»" ,data[i][byhand_column]);

    let tmp_table = tmp_body_1.getTables()[0];
    let students_list = get_students(data[i][small_bag_serial_column]);
    let students_table = tmp_table.getCell(1,0).setPaddingTop(0).setPaddingBottom(0).appendTable(students_list);


    let style = {}
    if (students_list.length > 23){
      style[DocumentApp.Attribute.FONT_SIZE] = 7;
    } else {
      style[DocumentApp.Attribute.FONT_SIZE] = 8;
    }
    students_table.setAttributes(style);
    for (j=0; j < students_table.getNumRows() ; j++){
      for (k=0; k < students_table.getRow(j).getNumCells(); k++){
        let cell_style = {}
        cell_style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
        students_table.getRow(j).getCell(k)
          .setPaddingTop(0)
          .setPaddingBottom(0)
          .getChild(0).asParagraph().setAttributes(cell_style);

        switch (k){
          case 0:  // 編號
            students_table.getRow(j).getCell(k).setWidth(20);
            break;
          case 1:  // 班級
            students_table.getRow(j).getCell(k).setWidth(60);
            break;
          case 2:  // 學號
            students_table.getRow(j).getCell(k).setWidth(60);
            break;
          case 3:  // 姓名
            students_table.getRow(j).getCell(k).setWidth(60);
            break;
          case 4:  // 科目名稱
            students_table.getRow(j).getCell(k).setWidth(140);
            break;
        }
      }
    }

    body.appendTable(tmp_table.copy());
    if (i % BATCH_SIZE != BATCH_SIZE - 1){
      body.appendPageBreak();
    }
    tmp_body_1 = tmp_body.copy();
    
    if(i % BATCH_SIZE == BATCH_SIZE - 1){
      // 設定字體後存檔
      body.editAsText().setFontFamily("Noto Sans TC");
      merged_doc.saveAndClose();

      // 轉換成 PDF 檔，且在檔名加上小袋序號的範圍
      let pdf_blob = merged_file.getAs('application/pdf');
      let pdf_name = merged_filename + "_" + ( 1 + BATCH_SIZE * Math.floor(i / BATCH_SIZE)).toLocaleString('en-US', {minimumIntegerDigits: number_of_digits, useGrouping:false}) + "-" + (BATCH_SIZE * (1 + Math.floor(i / BATCH_SIZE))).toLocaleString('en-US', {minimumIntegerDigits: number_of_digits, useGrouping:false})  + ".pdf";
      let pdf_file = DriveApp.createFile(pdf_blob).setName(pdf_name)
      pdf_file.moveTo(folder);

      // 把舊的 merged_file 移到垃圾桶後，再重新產生新的 merged_file
      merged_file.setTrashed(true);
      merged_file = small_template.makeCopy(merged_filename, folder);
      merged_doc = DocumentApp.openById(merged_file.getId());
      body = merged_doc.getBody();
      body.removeChild(body.getTables()[0]);
    }
  }

  // 將暫存檔移到垃圾桶
  tmp_doc.saveAndClose();
  DriveApp.getFileById(tmp_doc.getId()).setTrashed(true);

  if (data.length % BATCH_SIZE != 0){
    // 設定字體為 Noto Sans TC 後存檔
    body.editAsText().setFontFamily("Noto Sans TC");
    merged_doc.saveAndClose();

    // 把最後一批轉換成 PDF 檔
    pdf_blob = merged_file.getAs('application/pdf');
    pdf_name = merged_filename + "_" + (BATCH_SIZE * Math.floor(data.length / BATCH_SIZE).toLocaleString('en-US', {minimumIntegerDigits: number_of_digits, useGrouping:false}) + 1) + "-" + data.length + ".pdf";
    pdf_file = DriveApp.createFile(pdf_blob).setName(pdf_name)
    pdf_file.moveTo(folder);
  }
  merged_file.setTrashed(true);

  newRuntime = runtime_count_stop(runtime_count_start);
  SpreadsheetApp.getUi().alert("已完成小袋封面合併列印，共" + data.length + "頁，使用" + newRuntime + "秒。");
}


function get_students(small_bag_serial){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const [headers, ...data] = filtered_sheet.getDataRange().getValues();
  const small_bag_serial_column = headers.indexOf("小袋序號");
  const class_column = headers.indexOf("班級");
  const std_number_column = headers.indexOf("學號");
  const std_name_column = headers.indexOf("姓名");
  const subject_name_column = headers.indexOf("科目名稱");

  let students_table = [["", "班級", "學號", "姓名", "科目", "缺考"]]
  data.forEach(
    function (row, i){
      if (row[small_bag_serial_column] == small_bag_serial){
        students_table.push(
          [parseInt(students_table.length).toString() ,row[class_column], row[std_number_column], row[std_name_column], row[subject_name_column], ""]
        );
      }
    }
  );

  return students_table;
}
