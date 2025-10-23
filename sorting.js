function sort_by_subject(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filtered_sheet.getDataRange();

  filtered.offset(1,0,filtered.getNumRows()-1).sort(
    [
      {column: 2, ascending: true}, 
      {column: 3, ascending: true},
      {column: 9, ascending: true},
      {column: 10, ascending: true}, 
      {column: 8, ascending: true}, 
      {column: 5, ascending: true}      
    ]
  );
}


function sort_by_classname(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filtered_sheet.getDataRange();

  filtered.offset(1,0,filtered.getNumRows()-1).sort(
    [
      {column: 2, ascending: true}, 
      {column: 3, ascending: true}, 
      {column: 5, ascending: true}, 
      {column: 9, ascending: true}, 
      {column: 8, ascending: true}
    ]
  );
}


function sort_by_session_classroom(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filtered_sheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filtered_sheet.getDataRange();

  filtered.offset(1,0,filtered.getNumRows()-1).sort(
    [
      {column: 9, ascending: true},
      {column: 10, ascending: true},
      {column: 2, ascending: true},
      {column: 3, ascending: true},
      {column: 5, ascending: true},
      {column: 8, ascending: true},
    ]
  );
}
