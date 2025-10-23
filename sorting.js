function sortBySubject(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filteredSheet.getDataRange();

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


function sortByClassname(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filteredSheet.getDataRange();

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


function sortBySessionClassroom(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const filteredSheet = ss.getSheetByName("排入考程的補考名單");
  const filtered = filteredSheet.getDataRange();

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
