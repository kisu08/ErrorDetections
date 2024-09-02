function searchDataE2022() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 別のスプレッドシートのIDを指定
  var externalSpreadsheetId = '1HFVtazHKmEhlCTH3r2qAVJHFmPJaGKZkGwBU1Yr8Me4';
  
  // 別のスプレッドシートを開く
  var externalSpreadsheet = SpreadsheetApp.openById(externalSpreadsheetId);
  
  // シート名を確認
  var sheetAName = 'E（日経＋QUICK合計）';
  var sheetBName = '確認2022(日経E)';
  
  var sheetA = externalSpreadsheet.getSheetByName(sheetAName);
  var sheetB = ss.getSheetByName(sheetBName);

  // シートが正しく取得できているか確認
  if (!sheetA) {
    Logger.log('E（日経＋QUICK合計）が見つかりません: ' + sheetAName);
    SpreadsheetApp.getUi().alert('E（日経＋QUICK合計）が見つかりません: ' + sheetAName);
    return;
  }
  if (!sheetB) {
    Logger.log('確認2022(日経E)が見つかりません: ' + sheetBName);
    SpreadsheetApp.getUi().alert('確認2022(日経E)が見つかりません: ' + sheetBName);
    return;
  }

  // 確認2022(日経E)のB2セルの値を取得
  var codeToSearch = sheetB.getRange('B2').getValue();

  // E（日経＋QUICK合計）の範囲を取得
  var dataA = sheetA.getDataRange().getValues();

  // E（日経＋QUICK合計）のヘッダー行を取得
  var headersA = sheetA.getRange(3, 1, 1, sheetA.getLastColumn()).getValues()[0];

  // 銘柄コードの列インデックスを取得
  var codeColumnIndex = headersA.indexOf('コード');
  if (codeColumnIndex === -1) {
    SpreadsheetApp.getUi().alert('E（日経＋QUICK合計）に銘柄コード列が見つかりません。');
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

  // 確認2022(日経E)の6行目の項目名を取得
  var headersB = sheetB.getRange(6, 1, 1, sheetB.getLastColumn()).getValues()[0];
  
  // E（日経＋QUICK合計）の3行目の項目名を取得
  var headersA3 = sheetA.getRange(3, 1, 1, sheetA.getLastColumn()).getValues()[0];
  
  // E（日経＋QUICK合計）の2行目の項目名を取得
  var headersA2 = sheetA.getRange(2, 1, 1, sheetA.getLastColumn()).getValues()[0];

  var startRowB = 7; // 確認2022(日経E)にデータを反映させる開始行
  for (var r = 0; r < matchingRows.length; r++) {
    var rowIndex = matchingRows[r];
    var dataToReflect = dataA[rowIndex];

    for (var j = 0; j < headersB.length; j++) {
      if (j < 10) { // J列まで
        // 確認2022(日経E)の5行目のJ列までとE（日経＋QUICK合計）3行目の項目名を比較
        var headerIndexA3 = headersA3.indexOf(headersB[j]);
        if (headerIndexA3 !== -1) {
          var cell = sheetB.getRange(startRowB + r, j+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect[headerIndexA3]);
        }
      } else {
        // 確認2022(日経E)の5行目のJ列以降とE（日経＋QUICK合計）2行目の項目名を比較
        var headerIndexA2 = headersA2.indexOf(headersB[j]);
        if (headerIndexA2 !== -1) {
          var cell = sheetB.getRange(startRowB + r,j+1);
          cell.setNumberFormat('@'); //書式をテキストに設定
          cell.setValue(dataToReflect[headerIndexA2]);
        }
      }
    }
  }
  SpreadsheetApp.getUi().alert('抽出処理が正常に完了しました');
}