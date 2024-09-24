function deleteDataG2021() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  //シート名を確認
  var sheetName = '確認2021(日経G)';
  var sheet = ss.getSheetByName(sheetName);


  //シートが正しく取得できているか確認
  if(!sheet){
    Logger.log('確認2021(日経G)が見つかりません:' + sheetName);
    SpreadsheetApp.getUi().alert('確認2021(日経G)が見つかりません:' + sheetName);
    return;
  }

  //データの削除を実行
  var LastRow = sheet.getLastRow();
  var LastColumn = sheet.getLastColumn();
  sheet.getRange(7,1,LastRow-5,LastColumn).clear();
  sheet.getRange(5,1,1,LastColumn).clear();
  SpreadsheetApp.getUi().alert('削除処理が正常に完了しました');
}
