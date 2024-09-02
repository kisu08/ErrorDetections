function searchDataGdetail2023() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
  var externalSpreadsheetId = '1_b5l6hPz50367J-LvWxu5nMgM2i2d-bkj59NDRSqNI0';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  
  // シート名を確認
  var sheetAName = 'G詳細';
  var sheetBName = '確認2023(詳細G)';
  
  var sheetA = externalSpreadsheet.getSheetByName(sheetAName);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA) {
    Logger.log('G詳細が見つかりません: ' + sheetAName);
    SpreadsheetApp.getUi().alert('G詳細が見つかりません: ' + sheetAName);
    return;
  }
  if (!sheetB) {
    Logger.log('確認2023(詳細G)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('確認2023(詳細G)が見つかりません: ' + sheetBName);
    return;
  }

  // 確認2023(詳細G)のB2セルの値を取得
  var codeToSearch = sheetB.getRange('B2').getValue();

  // G詳細の範囲を取得
  var dataA = sheetA.getDataRange().getValues();

  // G詳細のヘッダー行を取得
  var headersA = sheetA.getRange(2, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex = headersA.indexOf('コード');
  if (codeColumnIndex === -1) {
    SpreadsheetApp.getUi().alert('G詳細に銘柄コード列が見つかりません。');
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

  // 確認2023(詳細G)の6行目の項目名を取得
  var headersB = sheetB.getRange(6, 1, 1, sheetB.getLastColumn()).getValues()[0];
  
  var startRowB = 7; // 確認2023(詳細G)にデータを反映させる開始行
  for (var r = 0; r < matchingRows.length; r++) {
    var rowIndex = matchingRows[r];
    var dataToReflect = dataA[rowIndex];

    for (var j = 0; j < headersB.length; j++) {
      // 確認2023(詳細G)の6行目のB列までとG詳細2行目の項目名を比較
      var headerIndexA = headersA.indexOf(headersB[j]);
      if (headerIndexA !== -1) {
        var cell = sheetB.getRange(startRowB + r, j+1);
        cell.setNumberFormat('@'); //書式をテキストに設定
        cell.setValue(dataToReflect[headerIndexA]);
      }
    }
  }
  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}