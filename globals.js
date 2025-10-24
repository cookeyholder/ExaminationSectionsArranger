const ACTIVE_SPREADSHEET = SpreadsheetApp.getActiveSpreadsheet();

const PARAMETERS_SHEET = ACTIVE_SPREADSHEET.getSheetByName("參數區");
const UNFILTERED_MAKEUP_SHEET = ACTIVE_SPREADSHEET.getSheetByName("註冊組補考名單");
const TEACHING_SELECTION_SHEET = ACTIVE_SPREADSHEET.getSheetByName("教學組排入考程的科目");
const OPEN_COURSE_LOOKUP_SHEET = ACTIVE_SPREADSHEET.getSheetByName("開課資料(查詢任課教師用)");
const FILTERED_RESULT_SHEET = ACTIVE_SPREADSHEET.getSheetByName("排入考程的補考名單");
const SMALL_BAG_DATA_SHEET = ACTIVE_SPREADSHEET.getSheetByName("小袋封面套印用資料");
const BIG_BAG_DATA_SHEET = ACTIVE_SPREADSHEET.getSheetByName("大袋封面套印用資料");
const BULLETIN_OUTPUT_SHEET = ACTIVE_SPREADSHEET.getSheetByName("公告版補考場次");
const RECORD_OUTPUT_SHEET = ACTIVE_SPREADSHEET.getSheetByName("試場紀錄表(A表)");
const SESSION_TIME_REFERENCE_SHEET = ACTIVE_SPREADSHEET.getSheetByName("節次時間表");
