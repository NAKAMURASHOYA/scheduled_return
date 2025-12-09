function reset_check() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName("キッティング(詳細版)");

  sh.getRange(2,1,1,4).clearContent();
  sh.getRange(4,1,1,4).clearContent();
  sh.getRange(6,6,250,1).clearContent();

}
