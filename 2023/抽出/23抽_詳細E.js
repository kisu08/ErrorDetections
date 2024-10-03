function searchDataEdetail2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
  var externalSpreadsheetId = '171cd4G1arGMUmHHkMeZdMj8S8Lo5mTbi4R63PgjspYA';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  
  // シート名を確認
  var sheetAName = 'E詳細';
  var sheetBName = '確認2023(詳細E)';
  
  var sheetA = externalSpreadsheet.getSheetByName(sheetAName);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA) {
    Logger.log('E詳細が見つかりません: ' + sheetAName);
    SpreadsheetApp.getUi().alert('E詳細が見つかりません: ' + sheetAName);
    return;
  }
  if (!sheetB) {
    Logger.log('確認2023(詳細E)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('確認2023(詳細E)が見つかりません: ' + sheetBName);
    return;
  }

  // 確認2023(詳細E)のB2セルの値を取得
  var codeToSearch = sheetB.getRange('C2').getValue();

  // E詳細の範囲を取得
  var dataA = sheetA.getDataRange().getValues();

  // E詳細のヘッダー行を取得
  var headersA = sheetA.getRange(2, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex = headersA.indexOf('コード');
  if (codeColumnIndex === -1) {
    SpreadsheetApp.getUi().alert('E詳細に銘柄コード列が見つかりません。');
    return;
  }

  var matchingRows = [];
  for (var i = 1; i < dataA.length; i++) {
    if (dataA[i][codeColumnIndex] == codeToSearch) {
      matchingRows.push(i);
    }
  }

  if (matchingRows.length == 0) {
    // 該当する銘柄コードが見つからない場合
    SpreadsheetApp.getUi().alert('該当する銘柄コードが見つかりませんでした。');
    return;
  }

  // 確認2023(詳細E)の6行目の項目名を取得
  var headersB = sheetB.getRange(6, 1, 1, sheetB.getLastColumn()).getValues()[0];
  
  var startRowB = 7; // 確認2023(詳細E)にデータを反映させる開始行
  for (var r = 0; r < matchingRows.length; r++) {
    var rowIndex = matchingRows[r];
    var dataToReflect = dataA[rowIndex];

    for (var k = 0; k < headersB.length; k++) {
      // 確認2023(詳細E)の6行目のB列までとE詳細2行目の項目名を比較
      var headerIndexA = headersA.indexOf(headersB[k]);
      if (headerIndexA !== -1) {
        var cell = sheetB.getRange(startRowB + r, k+1);
        cell.setValue(dataToReflect[headerIndexA]);
      }
    }
  }
  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}