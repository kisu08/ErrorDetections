function checkDataE2021() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("確認2021(日経E)");
  var data = sheet.getDataRange().getValues();
  var headers = data[5]; // 6行目が項目名
  var flagRow = 4; // 5行目にフラグを立てる

  // ヘッダーインデックスを先頭に定義
  var headerIndices = {
    documentNameCol: headers.indexOf("資料名称"),
    disclosureYearCol: headers.indexOf("開示年度"),
    typeNameCol: headers.indexOf("種別名"),
    sourceTypeCol: headers.indexOf("出典種別"),
    codeCol: headers.indexOf("コード"),
    pastYearCol: headers.indexOf("過年度：年"),
    pastYearMonthCol: headers.indexOf("過年度：年月／単位（加工値）"),
    startCol: headers.indexOf("【環境】温室効果ガス（GHG）排出量（Scope1）"),
    environmentalReserveCol: headers.indexOf("【環境】（予備）"),
    disclosureDateCol: headers.indexOf("資料公表日")
  };
  // 色がついているセルをカウントしてA5に表示する関数
function countColoredCells(sheet) {
  // A7以降のA列の範囲を取得
  var range = sheet.getRange('A7:A');
  var backgrounds = range.getBackgrounds();
  
  var count = 0;
  
  // 色がついているセルをカウント
  for (var i = 0; i < backgrounds.length; i++) {
    if (backgrounds[i][0] !== '#ffffff' && backgrounds[i][0] !== '') {
      count++;
    }
  }
  
  // A5A6セルを連結し、エラー行数を表示
    var cell = sheet.getRange('A5:A6');
    if (count > 0) {
      // エラーがある場合の表示
      cell.merge() // A5A6を連結
          .setValue("エラー行数: " + count + "行")
          .setFontColor('red')          // 赤文字
          .setFontWeight('bold')        // 太字
          .setHorizontalAlignment('center') // 水平方向の中央揃え
          .setVerticalAlignment('middle')   // 垂直方向の中央揃え
          .setFontSize(13)              // フォントサイズを13に設定
          .setBackground('#FFCCCC');    // 明るい赤色
    } else {
      // エラーがない場合の表示
      cell.merge() // A5A6を連結
          .setValue("データは正常です") // メッセージを「データは正常です」に変更
          .setFontColor('blue')         // 青文字
          .setFontWeight('bold')        // 太字
          .setHorizontalAlignment('center') // 水平方向の中央揃え
          .setVerticalAlignment('middle')   // 垂直方向の中央揃え
          .setFontSize(13)              // フォントサイズを13に設定
          .setBackground('#CCE5FF');    // 明るい青色
    }
  }

  // エラー検知してイエローに変更する関数
  function markErrorCell(row, col) {
    sheet.getRange(row + 1, col + 1).setBackground("yellow");
    sheet.getRange(flagRow + 1, col + 1).setValue(1);
    sheet.getRange(row + 1, 1).setValue("入力に不備があります");  // A列にエラーメッセージをセット
    sheet.getRange(row + 1, 1).setBackground("orange");
  }

  //行のループ処理を関数化
  function iterateRows(data, startRow, callback) {
    for (var row = startRow; row < data.length; row++) {
      callback(row);
    }
  }

  //列のループ処理を関数にして共通化
  function iterateCols(data, row, startCol, callback) {
    for (var col = startCol; col < data[row].length; col++) {
      callback(col);
    }
  }

  // 項目特有のエラー検知条件を設定する
  var conditions = {
    "出典種別": function(value, row) {
      var documentType = data[row][headerIndices.documentNameCol];
      var disclosureYear = data[row][headerIndices.disclosureYearCol];
      var documentValueMap = {
        "有価証券報告書": "004",
        "コーポレートガバナンス報告書": "005",
        "企業HP": "006"
      };

      if (disclosureYear === 2021) {
        return value === "";
      } else {
        return value === documentValueMap[documentType];
      }
    },

    "種別名": function(value, row) {
      var validValues = ["資料開示", "開示データ", "加工データ", "単位", "URL", "ページ数", "対象範囲"];
      var isValid = validValues.includes(value);

      if (value === "資料開示" || 
          (["URL", "ページ数", "対象範囲"].includes(value) && 
           ["有価証券報告書", "コーポレートガバナンス報告書"].includes(data[row][headerIndices.documentNameCol]))) {
        iterateCols(data, row, headerIndices.startCol, function(col) {
          if (data[row][col] !== "") {
            markErrorCells(row, headerIndices.startCol, data[row].length);
          }
        });
      }

      return isValid;
    },

    "コード": function(value) {
      return /^[0-9]{4}$/.test(value) || /^[A-Za-z0-9]{4}$/.test(value);
    },

    "資料名称": function(value) {
      if (["有価証券報告書", "コーポレートガバナンス報告書"].includes(value) && 
          ["URL", "ページ数", "対象範囲"].includes(data[headerIndices.typeNameCol])) {
        return false;
      }
      return ["企業HP", "有価証券報告書", "コーポレートガバナンス報告書"].includes(value);
    },

    "資料公表日": function(value, row) {
      var documentType = data[row][headerIndices.documentNameCol];
      if (["有価証券報告書", "コーポレートガバナンス報告書"].includes(documentType) && value === "") {
        return false;
      }

      if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
        return true;
      }

      
      // 日付オブジェクトかどうかを確認
      if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
        return true;
      }
      
      // 文字列として日付フォーマットを検証
      var valueStr = value.toString().trim();
      
      // 日付フォーマットを検証
      var datePattern = /^(?:(19|20)?\d\d年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|(?:19|20)?\d\d[-\/.](0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])|(?:0[1-9]|[12][0-9]|3[01])[-\/.](0[1-9]|1[0-2])[-\/.](19|20)?\d\d|(?:0[1-9]|1[0-2])[-\/.](0[1-9]|[12][0-9]|3[01])[-\/.](19|20)?\d\d|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月(0[1-9]|[12][0-9]|3[01])日|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月[一二三四五六七八九十]{1,2}日|(?:19|20)?\d\d年(0[1-9]|1[0-2])月|令和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|平成[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|昭和[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|大正[一二三四五六七八九十]{1,2}年(0[1-9]|1[0-2])月|[一二三四五六七八九十百千万]{1,4}年[一二三四五六七八九十]{1,2}月)|$/

      return datePattern.test(valueStr);
    },

    "開示年度": function(value) {
      return value === 2021;
    },

    "過年度：年": function(value, row) {
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      return /^[0-9]{4}$/.test(value);
    },

    "過年度：年月／単位（加工値）": function(value, row) {
      if (data[row][headerIndices.typeNameCol] === "資料開示") {
        return value === "";
      }
      var valueStr = value.toString();
      return /^[0-9]{6}$/.test(valueStr) && /^(00|01|02|03|04|05|06|07|08|09|10|11|12)$/.test(valueStr.slice(-2));
    }
  };

 // エラー検知とフラグ設定
 iterateRows(data, 6, function(row) {
  for (var col = 0; col < headers.length; col++) {
    var header = headers[col];
    var value = data[row][col];
    if (conditions[header] && !conditions[header](value, row)) {
      markErrorCell(row, col);
    }
  }
});

// 開示年度と資料名称の一致チェック
checkYearNameConsistency();


  // 資料開示の一貫性チェック
  if (!checkDocumentDisclosure() || !checkCombinationConsistency()) {
    markErrorColumn("種別名", "red");
  }

  // エラーチェック完了メッセージ
  SpreadsheetApp.getUi().alert('確認処理が正常に完了しました');

  // エラーレンジのマーク付け関数
  function markErrorCells(row, startCol, endCol) {
    for (var col = startCol; col < endCol; col++) {
      markErrorCell(row, col);
    }
  }

  // エラー列のマーク付け関数
  function markErrorColumn(headerName, color) {
    var colIndex = headers.indexOf(headerName);
    sheet.getRange(flagRow + 1, colIndex + 1).setBackground(color);
  }

  // 開示年度と資料名称の一致チェック関数
  function checkYearNameConsistency() {
    var dataMap = {};
    var yearIndex = headerIndices.disclosureYearCol;
    var nameIndex = headerIndices.documentNameCol;
    var dateIndex = headerIndices.disclosureDateCol;

    for (var row = 6; row < data.length; row++) {
      var year = data[row][yearIndex];
      var name = data[row][nameIndex].trim();
      var date = data[row][dateIndex].toString().trim();
      var key = year + "_" + name;
      if (!dataMap[key]) {
        dataMap[key] = date;
      } else if (dataMap[key] !== date) {
        markErrorCell(row, dateIndex);
      }
    }
  }

  // 同一の開示年度では、1資料につき必ず資料開示のレコードは1つであること
  function checkDocumentDisclosure() {
    var disclosureCount = 0;
    var uniqueCombinations = new Set();

    iterateRows(data, 6, function(row) {
      var documentName = data[row][headerIndices.documentNameCol];
      var typeName = data[row][headerIndices.typeNameCol];
      var disclosureYear = data[row][headerIndices.disclosureYearCol];

      if (typeName === "資料開示") {
        disclosureCount++;
      }

      var combination = documentName + "_" + disclosureYear;
      uniqueCombinations.add(combination);
    });

    return uniqueCombinations.size === disclosureCount;
  }

  // 種別名ごとの一貫性チェック関数
  function checkCombinationConsistency() {
    var baseCombinations = {};

    iterateRows(data, 6, function(row) {
      var typeName = data[row][headerIndices.typeNameCol];
      var documentName = data[row][headerIndices.documentNameCol];
      var pastYear = data[row][headerIndices.pastYearCol];
      var pastYearMonth = data[row][headerIndices.pastYearMonthCol];

      if (typeName === "資料開示") return;

      var combination = pastYear + "_" + pastYearMonth;
      if (!baseCombinations[documentName]) {
        baseCombinations[documentName] = {};
      }

      if (!baseCombinations[documentName][typeName]) {
        baseCombinations[documentName][typeName] = new Set();
      }
      baseCombinations[documentName][typeName].add(combination);
    });

    for (var documentName in baseCombinations) {
      var typeNameSets = baseCombinations[documentName];
      var allSets = new Set();
      for (var typeName in typeNameSets) {
        if (allSets.size === 0) {
          allSets = new Set(typeNameSets[typeName]);
        } else {
          if (allSets.size !== typeNameSets[typeName].size || 
              ![...allSets].every(value => typeNameSets[typeName].has(value))) {
            console.log(`Mismatch found in document: ${documentName}, type: ${typeName}`);
            sheet.getRange(flagRow + 1, headerIndices.typeNameCol + 1).setValue(1);
            return false;
          }
        }
      }
    }
    return true;
  }
  
   // エラー行のカウントを実行
   countColoredCells(sheet);
}